"""将系统生成的 PPTX 转成同一版本的 PDF 和逐页 PNG。"""

import hashlib
import json
import os
import re
import shutil
import signal
import subprocess
import tempfile
import zipfile
from pathlib import Path

from PIL import Image


class PreviewError(RuntimeError):
    pass


def _run(arguments, timeout=120):
    environment = {key: value for key, value in os.environ.items()
                   if not any(word in key.upper() for word in ("API_KEY", "TOKEN", "SECRET", "PASSWORD"))
                   and key != "DASHSCOPE_BASE_URL"}
    environment.update(LANG="C", LC_ALL="C")
    options = {"creationflags": subprocess.CREATE_NO_WINDOW | subprocess.CREATE_NEW_PROCESS_GROUP} if os.name == "nt" else {"start_new_session": True}
    process = subprocess.Popen(arguments, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=environment, **options)
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired as error:
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"], capture_output=True,
                           timeout=10, creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            os.killpg(process.pid, signal.SIGKILL)
        process.communicate(timeout=10)
        raise PreviewError("成品渲染超时，页面保持未审核") from error
    if process.returncode:
        raise PreviewError(f"成品渲染工具执行失败（退出码 {process.returncode}），页面保持未审核")
    return (stdout + stderr).decode("utf-8", errors="replace")


def discover_tools():
    project = Path(__file__).resolve().parents[1]
    settings_path = project / ".runtime/preview-tools.json"
    settings = json.loads(settings_path.read_text(encoding="utf-8")) if settings_path.is_file() else {}
    office_candidates = [os.environ.get("PPT_SOFFICE"), shutil.which("soffice.com"), shutil.which("soffice"),
                         project / ".runtime/libreoffice/program/soffice.com",
                         project / ".runtime/libreoffice/program/soffice.exe",
                         Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "LibreOffice/program/soffice.com"]
    office = next((str(Path(item).resolve()) for item in office_candidates if item and Path(item).is_file()), None)
    poppler = os.environ.get("PPT_POPPLER_BIN") or settings.get("poppler_bin")
    result = {"soffice": office}
    for name in ("pdfinfo", "pdftoppm"):
        candidate = Path(poppler) / (name + (".exe" if os.name == "nt" else "")) if poppler else None
        result[name] = str(candidate.resolve()) if candidate and candidate.is_file() else shutil.which(name)
    return result


def tool_signature(tools=None):
    tools = tools or discover_tools()
    missing = [name for name, value in tools.items() if not value]
    if missing:
        raise PreviewError("成品预览缺少运行时：" + "、".join(missing) + "；PPTX 可下载，页面保持未审核")
    versions = {name: _run([value, "--version" if name == "soffice" else "-v"], timeout=15).strip().splitlines()[0][:180]
                for name, value in tools.items()}
    font_files = []
    if os.name == "nt":
        directory = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
        font_files = [directory / name for name in ("msyh.ttc", "msyhbd.ttc", "simhei.ttf", "arial.ttf", "seguiemj.ttf")]
    elif shutil.which("fc-match"):
        font_files = [Path(_run(["fc-match", "-f", "%{file}", family], timeout=10).strip())
                      for family in ("Microsoft YaHei", "Noto Sans CJK SC", "Arial")]
    fonts = {file.name: hashlib.sha256(file.read_bytes()).hexdigest() for file in font_files if file.is_file()}
    return {"versions": versions, "font_files_sha256": fonts,
            "pipeline_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
            "font_override": os.environ.get("PPT_FONT_SIGNATURE", ""), "pixel_limit": 1600}


def _check_input(path):
    if path.suffix.lower() != ".pptx" or path.stat().st_size > 30 * 1024 * 1024:
        raise PreviewError("预览只接受不超过 30 MiB 的系统生成 PPTX")
    with zipfile.ZipFile(path) as archive:
        if sum(item.file_size for item in archive.infolist()) > 200 * 1024 * 1024:
            raise PreviewError("PPTX 解压内容超出预览限制")
        if any("vbaproject" in name.lower() for name in archive.namelist()):
            raise PreviewError("预览不处理包含宏的文稿")
        for name in archive.namelist():
            if name.endswith(".rels") and re.search(rb'TargetMode\s*=\s*[\'"]External[\'"]', archive.read(name)):
                raise PreviewError("预览不处理包含外部资源引用的文稿")


def render_preview(filename, output_dir, signature=None, tools=None):
    """输入、工具与字体绑定缓存；每次转换使用独立 LibreOffice 配置。"""
    path = Path(filename).resolve()
    tools = tools or discover_tools()
    signature = signature or tool_signature(tools)
    _check_input(path)
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    identity = hashlib.sha256(json.dumps({"pptx": digest, "runtime": signature}, sort_keys=True).encode()).hexdigest()
    parent = Path(output_dir).resolve()
    parent.mkdir(parents=True, exist_ok=True)
    target = parent / identity[:24]
    metadata = target / "preview.json"
    if metadata.is_file():
        saved = json.loads(metadata.read_text(encoding="utf-8"))
        if saved.get("cache_key") == identity and all((target / item["name"]).is_file() and
                hashlib.sha256((target / item["name"]).read_bytes()).hexdigest() == item["sha256"] for item in saved["files"]):
            return {**saved, "pdf": str(target / "deck.pdf"),
                    "pages": [str(target / item["name"]) for item in saved["files"] if item["role"] == "preview"]}
        raise PreviewError("预览缓存资产已改变，未复用或覆盖旧截图")
    if target.exists():
        raise PreviewError("预览目标目录不完整，未覆盖已有资产")
    scratch = Path(os.environ.get("PPT_RENDER_TMP") or Path(__file__).resolve().parents[1] / ".runtime/render-tmp").resolve()
    scratch.mkdir(parents=True, exist_ok=True)
    # Windows 原生转换器仍受长路径影响，用户配置目录必须使用短的独立路径。
    with tempfile.TemporaryDirectory(prefix="render-", dir=scratch) as temporary:
        build = Path(temporary)
        source = build / "deck.pptx"
        shutil.copyfile(path, source)
        if hashlib.sha256(source.read_bytes()).hexdigest() != digest:
            raise PreviewError("PPTX 在预览开始时已改变，未关联截图")
        profile = build / "profile"
        profile.mkdir()
        _run([tools["soffice"], "-env:UserInstallation=" + profile.as_uri(), "--headless", "--nologo", "--nodefault",
              "--norestore", "--convert-to", "pdf:impress_pdf_Export", "--outdir", str(build), str(source)])
        pdf = build / "deck.pdf"
        if not pdf.is_file() or pdf.stat().st_size > 100 * 1024 * 1024:
            raise PreviewError("LibreOffice 未生成有效 PDF，页面保持未审核")
        information = _run([tools["pdfinfo"], str(pdf)], timeout=20)
        match = re.search(r"^Pages:\s+(\d+)", information, re.M)
        from pptx import Presentation
        expected = len(Presentation(path).slides)
        if not match or int(match.group(1)) != expected or not 1 <= expected <= 50:
            raise PreviewError("成品 PDF 页数与 PPTX 不一致，页面保持未审核")
        _run([tools["pdftoppm"], "-png", "-r", "120", "-scale-to", "1600", str(pdf), str(build / "page")])
        pages = sorted(build.glob("page-*.png"), key=lambda item: int(item.stem.split("-")[-1]))
        if len(pages) != expected:
            raise PreviewError("成品截图页数不完整，页面保持未审核")
        if sum(item.stat().st_size for item in pages) > 100 * 1024 * 1024:
            raise PreviewError("成品截图超过容量限制，页面保持未审核")
        for image in pages:
            with Image.open(image) as picture:
                if picture.width * picture.height > 4_000_000:
                    raise PreviewError("成品截图像素超出限制")
                picture.verify()
        with tempfile.TemporaryDirectory(prefix=".publish-", dir=parent) as publish_directory:
            publish = Path(publish_directory) / "complete"
            publish.mkdir()
            files = []
            for file, role in [(pdf, "pdf"), *[(page, "preview") for page in pages]]:
                destination = publish / file.name
                shutil.copyfile(file, destination)
                files.append({"name": file.name, "role": role, "sha256": hashlib.sha256(destination.read_bytes()).hexdigest()})
            saved = {"cache_key": identity, "pptx_sha256": digest, "runtime": signature, "files": files, "page_count": expected}
            (publish / "preview.json").write_text(json.dumps(saved, ensure_ascii=False, indent=2), encoding="utf-8")
            if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
                raise PreviewError("PPTX 在预览期间已改变，未提交截图")
            os.replace(publish, target)
    return {**saved, "pdf": str(target / "deck.pdf"), "pages": [str(target / page.name) for page in pages]}
