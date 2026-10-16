"""持久任务 worker 与只读归档扫描、备份恢复入口。"""

import argparse
import json
import logging
from pathlib import Path

from config.archives import create_backup, restore_backup, scan_archives
from config.jobs import dispatch_pending
from config.storage import PPTStore
from config_manager import get_output_root


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, help="与 Web 相同的输出根目录")
    parser.add_argument("--once", action="store_true", help="执行一条已排队任务后退出，供诊断使用")
    parser.add_argument("--scan-archives", action="store_true", help="只读扫描旧归档，不导入")
    parser.add_argument("--backup", type=Path, help="备份到输出根目录以外的新目录")
    parser.add_argument("--restore", type=Path, help="从备份恢复到 --root 指定的新目录")
    parser.add_argument("--preview-tools", action="store_true", help="检查成品预览工具与字体，不执行模型调用")
    parser.add_argument("--check-runtime", action="store_true", help="检查渲染模块与 Cairo，不消费任务或调用模型")
    args = parser.parse_args()
    if args.check_runtime:
        import sys
        from config_manager import RenderDependencyError, configure_cairo_runtime
        try:
            from renderer import cairosvg
            image = cairosvg.svg2png(bytestring=b'<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2"><rect width="2" height="2" fill="blue"/></svg>')
            print(json.dumps({"status": "ready", "python": sys.executable,
                              "cairo_directory": configure_cairo_runtime(), "svg_png_bytes": len(image)}, ensure_ascii=False, indent=2))
        except RenderDependencyError as error:
            print(json.dumps({"status": "failed", "error": str(error)}, ensure_ascii=False, indent=2))
            raise SystemExit(1)
        return
    if args.preview_tools:
        from config.preview import PreviewError, discover_tools, tool_signature
        try:
            print(json.dumps(tool_signature(), ensure_ascii=False, indent=2))
        except PreviewError as error:
            print(json.dumps({"tools": discover_tools(), "error": str(error)}, ensure_ascii=False, indent=2))
        return
    if args.restore:
        if not args.root:
            parser.error("恢复时必须明确指定 --root 新目录")
        print(restore_backup(args.restore, args.root))
        return
    root = (args.root or get_output_root()).resolve()
    store = PPTStore(root)
    if args.scan_archives:
        print(json.dumps(scan_archives(store), ensure_ascii=False, indent=2))
        return
    if args.backup:
        print(create_backup(store, args.backup))
        return
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    # HTTP 库的普通请求日志会带出模型服务地址，任务进度另有持久记录。
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.info("Worker output root: %s", root)
    queue = dispatch_pending(root, recover=True)
    try:
        if args.once:
            task = queue.dequeue()
            if task:
                queue.execute(task)
            else:
                print("没有待执行任务")
        else:
            # 单进程单执行线程，避免 SDK 全局配置与本地密钥环境发生竞争。
            queue.create_consumer(workers=1, worker_type="thread", periodic=True).run()
    finally:
        queue.storage.close()


if __name__ == "__main__":
    main()
