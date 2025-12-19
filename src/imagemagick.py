from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class ImageMagickConfig:
    command: list[str]
    env: dict[str, str] | None
    source: str


def _env_with_tools_dir(tools_dir: Path) -> dict[str, str]:
    tools_dir_str = str(tools_dir)
    env = dict(os.environ)
    env.setdefault("MAGICK_HOME", tools_dir_str)
    env.setdefault("MAGICK_CONFIGURE_PATH", tools_dir_str)
    env["PATH"] = tools_dir_str + os.pathsep + env.get("PATH", "")
    return env


def _looks_like_path(value: str) -> bool:
    return any(sep in value for sep in ("/", "\\")) or value.lower().endswith(".exe")


def resolve_imagemagick(base_dir: Path | None = None) -> ImageMagickConfig | None:
    """
    Resolve an ImageMagick CLI entrypoint in a cross-platform way.

    - Override via IMAGEMAGICK_BIN / MAGICK_BIN
    - Windows: prefer bundled `src/tools/magick.exe` if present
    - Linux/macOS: prefer `magick`, then `convert` (ImageMagick 6)
    """

    override = (os.getenv("IMAGEMAGICK_BIN") or os.getenv("MAGICK_BIN") or "").strip()
    if override:
        override_path = Path(override) if _looks_like_path(override) else None
        if override_path and override_path.exists():
            if override_path.is_dir():
                system = platform.system().lower()
                candidates = (
                    ["magick.exe", "magick"]
                    if system == "windows"
                    else ["magick", "convert"]
                )
                for name in candidates:
                    exe = override_path / name
                    if exe.exists():
                        return ImageMagickConfig(
                            command=[str(exe)],
                            env=_env_with_tools_dir(override_path),
                            source="env:dir",
                        )
                return None

            return ImageMagickConfig(
                command=[str(override_path)],
                env=_env_with_tools_dir(override_path.parent),
                source="env:file",
            )

        found = shutil.which(override)
        if found:
            return ImageMagickConfig(command=[found], env=None, source="env:which")
        return ImageMagickConfig(command=[override], env=None, source="env:raw")

    system = platform.system().lower()

    if system == "windows":
        if base_dir is not None:
            bundled = base_dir / "tools" / "magick.exe"
            if bundled.exists():
                env = _env_with_tools_dir(bundled.parent)
                return ImageMagickConfig(command=[str(bundled)], env=env, source="bundled")

        magick = shutil.which("magick")
        if magick:
            return ImageMagickConfig(command=[magick], env=None, source="path:magick")

        # Avoid `convert` on Windows: it often resolves to a built-in tool.
        return None

    if base_dir is not None:
        bundled = base_dir / "tools" / "magick"
        if bundled.exists():
            env = _env_with_tools_dir(bundled.parent)
            return ImageMagickConfig(command=[str(bundled)], env=env, source="bundled")

    magick = shutil.which("magick")
    if magick:
        return ImageMagickConfig(command=[magick], env=None, source="path:magick")

    convert = shutil.which("convert")
    if convert:
        return ImageMagickConfig(command=[convert], env=None, source="path:convert")

    return None


def _run_command(
    args: list[str],
    *,
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
    timeout_s: int | None = None,
) -> subprocess.CompletedProcess[str]:
    try:
        return subprocess.run(
            args,
            cwd=str(cwd) if cwd else None,
            env=env,
            check=True,
            capture_output=True,
            text=True,
            timeout=timeout_s,
        )
    except FileNotFoundError as e:
        raise RuntimeError(f"Command not found: {args[0]}") from e
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"Command timed out: {' '.join(args)}") from e
    except subprocess.CalledProcessError as e:
        stdout = (e.stdout or "").strip()
        stderr = (e.stderr or "").strip()
        detail = "\n".join(
            part
            for part in (
                f"Command failed: {' '.join(args)}",
                f"Exit code: {e.returncode}",
                f"stdout: {stdout}" if stdout else "",
                f"stderr: {stderr}" if stderr else "",
            )
            if part
        )
        raise RuntimeError(detail) from e


def _is_up_to_date(source: Path, target: Path) -> bool:
    if not target.exists():
        return False
    try:
        return target.stat().st_mtime >= source.stat().st_mtime and target.stat().st_size > 0
    except FileNotFoundError:
        return False


def _require_nonempty_file(path: Path, *, context: str) -> None:
    if not path.exists():
        raise RuntimeError(f"{context}: output file not found: {path}")
    try:
        if path.stat().st_size <= 0:
            raise RuntimeError(f"{context}: output file is empty: {path}")
    except FileNotFoundError as e:
        raise RuntimeError(f"{context}: output file not found: {path}") from e


def _atomic_replace(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    os.replace(src, dst)


_MAGICK_FORMAT_CACHE: dict[str, bool] = {}


def _magick_supports_format(format_name: str) -> bool:
    # Cache across calls; `magick -list format` can be expensive.
    key = format_name.strip().upper()
    cached = _MAGICK_FORMAT_CACHE.get(key)
    if cached is not None:
        return cached

    cp = _run_command(["magick", "-list", "format"])
    supported = False
    for line in (cp.stdout or "").splitlines():
        # Example: "     WEBP* rw+   WebP Image Format (libwebp 1.2.4)"
        if line.lstrip().upper().startswith(key):
            supported = True
            break
    _MAGICK_FORMAT_CACHE[key] = supported
    return supported


def _imagemagick_identify(path: Path, fmt: str) -> str:
    cp = _run_command(["magick", "identify", "-format", fmt, str(path)])
    return (cp.stdout or "").strip()


def _try_get_int(value: str) -> int | None:
    try:
        return int(value.strip())
    except Exception:
        return None


def _should_use_lossless_webp(raster_path: Path) -> bool:
    # Heuristic for "text / charts" images:
    # - Palette/Grayscale/Bilevel types are likely diagrams/screenshots.
    # - Non-opaque (alpha) often indicates UI assets/screenshots.
    # - Low unique color count suggests diagrams.
    try:
        img_type = _imagemagick_identify(raster_path, "%[type]")
        opaque = _imagemagick_identify(raster_path, "%[opaque]").lower()
        if img_type in {"Bilevel", "Grayscale", "GrayscaleAlpha", "Palette", "PaletteAlpha"}:
            return True
        if opaque in {"false", "0", "no"}:
            return True

        # `%k` (unique colors) may be expensive; guard with file size.
        try:
            if raster_path.stat().st_size <= 20 * 1024 * 1024:
                colors_raw = _imagemagick_identify(raster_path, "%k")
                colors = _try_get_int(colors_raw)
                if colors is not None and colors <= 512:
                    return True
        except FileNotFoundError:
            return False
    except Exception:
        # Fall back to a conservative default (lossy) to avoid blocking the pipeline.
        return False

    return False


def _iter_image_files(root: Path) -> Iterable[Path]:
    for dirpath, _, filenames in os.walk(root):
        base = Path(dirpath)
        for name in filenames:
            yield base / name


def _compute_collision_stems(paths: Iterable[Path]) -> set[tuple[Path, str]]:
    # Returns (parent_dir, stem_lower) pairs that appear multiple times with different extensions.
    counts: dict[tuple[Path, str], int] = {}
    for p in paths:
        key = (p.parent, p.stem.lower())
        counts[key] = counts.get(key, 0) + 1
    return {key for key, count in counts.items() if count > 1}


def _raster_output_path(src: Path, collision_stems: set[tuple[Path, str]]) -> Path:
    if (src.parent, src.stem.lower()) in collision_stems:
        # Stable, collision-free naming: image.png -> image_png.webp
        ext = src.suffix.lower().lstrip(".") or "img"
        return src.with_name(f"{src.stem}_{ext}.webp")
    return src.with_suffix(".webp")


def _convert_raster_to_webp(src: Path, dst: Path) -> None:
    if _is_up_to_date(src, dst):
        return

    lossless = _should_use_lossless_webp(src)
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        tmp_out = tmpdir_path / dst.name

        args = ["magick", str(src), "-auto-orient", "-strip"]
        if lossless:
            args += [
                "-define",
                "webp:lossless=true",
                "-define",
                "webp:method=6",
                "-define",
                "webp:image-hint=graph",
                str(tmp_out),
            ]
        else:
            args += [
                "-quality",
                "82",
                "-define",
                "webp:method=6",
                "-define",
                "webp:image-hint=photo",
                str(tmp_out),
            ]

        _run_command(args)
        _require_nonempty_file(tmp_out, context=f"Raster→WebP ({src.name})")
        _atomic_replace(tmp_out, dst)


def _libreoffice_convert(
    src: Path,
    *,
    convert_to: str,
    outdir: Path,
    timeout_s: int = 90,
) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as profile_dir:
        profile_path = Path(profile_dir)
        profile_uri = "file://" + profile_path.as_posix()

        _run_command(
            [
                "libreoffice",
                "--headless",
                "--nologo",
                "--nodefault",
                "--nolockcheck",
                "--norestore",
                "--invisible",
                f"-env:UserInstallation={profile_uri}",
                "--convert-to",
                convert_to,
                "--outdir",
                str(outdir),
                str(src),
            ],
            timeout_s=timeout_s,
        )

    # LibreOffice chooses output name; search for the newest matching extension.
    target_ext = "." + convert_to.split(":")[0].lower()
    candidates = list(outdir.glob(f"{src.stem}*{target_ext}"))
    if not candidates:
        # Some LibreOffice builds keep original extension in the name; fall back to scanning by ext.
        candidates = list(outdir.glob(f"*{target_ext}"))

    if not candidates:
        raise RuntimeError(f"LibreOffice did not produce {target_ext} from {src.name}")

    newest = max(candidates, key=lambda p: p.stat().st_mtime)
    _require_nonempty_file(newest, context=f"LibreOffice→{target_ext} ({src.name})")
    return newest


def _wmf_to_svg(wmf_path: Path, svg_path: Path) -> bool:
    if _is_up_to_date(wmf_path, svg_path):
        return True

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        for spec in ("svg", "svg:draw_svg_Export"):
            try:
                produced = _libreoffice_convert(wmf_path, convert_to=spec, outdir=tmpdir_path)
                tmp_svg = tmpdir_path / svg_path.name
                if produced != tmp_svg:
                    _atomic_replace(produced, tmp_svg)
                _require_nonempty_file(tmp_svg, context=f"WMF→SVG ({wmf_path.name})")
                _atomic_replace(tmp_svg, svg_path)
                return True
            except Exception:
                continue

    return False


def _pdf_to_raster(
    pdf_path: Path,
    out_path: Path,
    *,
    prefer_webp: bool,
) -> None:
    if _is_up_to_date(pdf_path, out_path):
        return

    if prefer_webp and not _magick_supports_format("WEBP"):
        prefer_webp = False

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        tmp_out = tmpdir_path / out_path.name

        # Force first page to guarantee a single output file.
        pdf_input = f"{pdf_path}[0]"
        common = [
            "magick",
            "-density",
            "300",
            pdf_input,
            "-background",
            "white",
            "-alpha",
            "remove",
            "-alpha",
            "off",
            "-strip",
        ]

        if prefer_webp:
            args = common + [
                "-define",
                "webp:lossless=true",
                "-define",
                "webp:method=6",
                "-define",
                "webp:image-hint=graph",
                str(tmp_out),
            ]
        else:
            args = common + [str(tmp_out)]

        _run_command(args)
        _require_nonempty_file(tmp_out, context=f"PDF→Raster ({pdf_path.name})")
        _atomic_replace(tmp_out, out_path)


def _wmf_fallback_to_raster(wmf_path: Path, out_path: Path) -> None:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        produced_pdf = _libreoffice_convert(wmf_path, convert_to="pdf", outdir=tmpdir_path)

        prefer_webp = out_path.suffix.lower() == ".webp"
        _pdf_to_raster(produced_pdf, out_path, prefer_webp=prefer_webp)


def optimize_extracted_images(image_dir: str) -> None:
    # Optimize DOCX-extracted images for web delivery:
    # - Raster (PNG/JPG/JPEG) -> WebP (lossless for text/charts via heuristic)
    # - WMF -> SVG via LibreOffice; fallback WMF->PDF->WebP(lossless) or PNG
    root = Path(image_dir)
    if not root.exists():
        raise RuntimeError(f"image_dir does not exist: {root}")
    if not root.is_dir():
        raise RuntimeError(f"image_dir is not a directory: {root}")

    files = [p for p in _iter_image_files(root) if p.is_file()]
    raster_exts = {".png", ".jpg", ".jpeg"}
    wmf_exts = {".wmf"}

    raster_sources = [p for p in files if p.suffix.lower() in raster_exts]
    collision_stems = _compute_collision_stems(raster_sources)

    errors: list[str] = []

    # Raster first (WebP is typically what the web will use even when WMF falls back).
    for src in raster_sources:
        try:
            dst = _raster_output_path(src, collision_stems)
            _convert_raster_to_webp(src, dst)
            if dst.exists() and dst.stat().st_size > 0 and src.exists():
                try:
                    src.unlink()
                except Exception as e:
                    errors.append(f"Failed to delete source raster {src}: {e}")
        except Exception as e:
            errors.append(f"Raster optimize failed for {src}: {e}")

    # WMF: try to preserve vector via SVG; otherwise rasterize.
    for wmf in (p for p in files if p.suffix.lower() in wmf_exts):
        try:
            svg = wmf.with_suffix(".svg")
            if _wmf_to_svg(wmf, svg):
                if svg.exists() and svg.stat().st_size > 0 and wmf.exists():
                    try:
                        wmf.unlink()
                    except Exception as e:
                        errors.append(f"Failed to delete source WMF {wmf}: {e}")
                continue

            # Fallback: WMF -> PDF -> WebP(lossless) or PNG if WebP unsupported
            if _magick_supports_format("WEBP"):
                raster_out = wmf.with_suffix(".webp")
            else:
                raster_out = wmf.with_suffix(".png")

            if not _is_up_to_date(wmf, raster_out):
                _wmf_fallback_to_raster(wmf, raster_out)

            if raster_out.exists() and raster_out.stat().st_size > 0 and wmf.exists():
                try:
                    wmf.unlink()
                except Exception as e:
                    errors.append(f"Failed to delete source WMF {wmf}: {e}")
        except Exception as e:
            errors.append(f"WMF optimize failed for {wmf}: {e}")

    if errors:
        raise RuntimeError("optimize_extracted_images encountered errors:\n" + "\n".join(errors))
