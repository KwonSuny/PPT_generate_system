# -*- coding: utf-8 -*-
"""API Key DPAPI 加密存储测试：加密往返、明文迁移、解密失败降级、注销清空。"""
import base64
import json
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ.pop("DASHSCOPE_API_KEY", None)

import config_manager

_tmp = Path(tempfile.mkdtemp(prefix="ppt_encrypt_test_"))
config_manager.CONFIG_DIR = _tmp / "config"
config_manager.CONFIG_FILE = config_manager.CONFIG_DIR / "config.json"


def check(name, cond):
    print(("PASS: " if cond else "FAIL: ") + name)
    if not cond:
        raise SystemExit(1)


def read_cfg():
    return json.loads(config_manager.CONFIG_FILE.read_text(encoding="utf-8"))


# 1) 加密往返：save 后 load 还原原文，落盘无明文
KEY = "sk-encrypted-abcdef123456"
config_manager.save_api_key(KEY)
check("加密后 load_api_key 还原原文", config_manager.load_api_key() == KEY)
check("has_api_key=True", config_manager.has_api_key() is True)
raw = config_manager.CONFIG_FILE.read_text(encoding="utf-8")
check("config.json 不含明文 key", KEY not in raw)
cfg = read_cfg()
check("明文 api_key 字段被清空", cfg.get("api_key") == "")
check("密文 api_key_encrypted 非空", bool(cfg.get("api_key_encrypted")))

# 2) 解密失败降级：伪造密文 -> 返回空，不崩溃
config_manager.save_api_key("sk-temp-whatever")
cfg = read_cfg()
cfg["api_key_encrypted"] = base64.b64encode(b"not-a-valid-dpapi-blob").decode("ascii")
config_manager.save_config(cfg)
check("伪造密文解密失败返回空", config_manager.load_api_key() == "")
check("伪造密文下 has_api_key=False", config_manager.has_api_key() is False)

# 3) 旧明文自动迁移：写入明文 -> load 后加密并清除明文
KEY2 = "sk-legacy-plaintext-999"
config_manager.save_config({"api_key": KEY2, "api_key_encrypted": ""})
check("明文迁移后仍可读取", config_manager.load_api_key() == KEY2)
cfg = read_cfg()
check("迁移后明文被清除", cfg.get("api_key") == "")
check("迁移后生成密文", bool(cfg.get("api_key_encrypted")))
raw = config_manager.CONFIG_FILE.read_text(encoding="utf-8")
check("迁移后文件不含明文", KEY2 not in raw)

# 4) 注销清空：两个字段均置空
config_manager.save_api_key("sk-after-migrate")
config_manager.clear_api_key()
check("注销后 load 为空", config_manager.load_api_key() == "")
cfg = read_cfg()
check("注销后两个字段均空", cfg.get("api_key") == "" and cfg.get("api_key_encrypted") == "")

print("\nALL ENCRYPTION TESTS PASSED")
print(f"临时隔离目录: {_tmp}")
