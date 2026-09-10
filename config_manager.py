
"""应用配置与路径管理：API key 本地持久化 + 输出目录定位。

- API key 使用 Windows DPAPI（数据保护接口）加密后存储于 config/config.json，
  密文绑定当前 Windows 用户与机器：换账户、换机器、重装系统后无法解密，
  需重新输入一次。config.json 中不再出现明文 key。
- 任何日志 / 接口响应均不回显完整 key。
"""

import base64
import json
from pathlib import Path

CONFIG_DIR = Path(__file__).resolve().parent / "config"
CONFIG_FILE = CONFIG_DIR / "config.json"

# DPAPI 描述标识（非密钥，仅作用途标记）
_DPAPI_DESC = "PPT Generator API Key"
_CRYPTPROTECT_UI_FORBIDDEN = 0x01

try:
    import win32crypt  # noqa: F401
    _DPAPI_AVAILABLE = True
except Exception:  # 非 Windows 或未安装 pywin32
    _DPAPI_AVAILABLE = False


def _default_config() -> dict:
    return {"api_key": "", "api_key_encrypted": "", "base_url": ""}


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        return _default_config()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return _default_config()
        cfg = _default_config()
        cfg.update(data)
        return cfg
    except Exception:
        return _default_config()


def save_config(cfg: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _encrypt(key: str) -> str:
    """用 DPAPI 加密明文 key，返回 base64 编码的密文字符串。"""
    if not _DPAPI_AVAILABLE:
        raise RuntimeError("当前环境不支持 DPAPI 加密（需 Windows 且安装 pywin32）")
    try:
        cipher = win32crypt.CryptProtectData(
            key.encode("utf-8"),
            _DPAPI_DESC,
            None,
            None,
            None,
            _CRYPTPROTECT_UI_FORBIDDEN,
        )
    except Exception as e:
        raise RuntimeError(f"DPAPI 加密失败: {e}") from e
    return base64.b64encode(cipher).decode("ascii")


def _decrypt(encoded: str) -> str:
    """解密 base64 密文，返回明文 key。换机/换账户/重装后抛异常。"""
    if not _DPAPI_AVAILABLE:
        raise RuntimeError("当前环境不支持 DPAPI 解密（需 Windows 且安装 pywin32）")
    try:
        cipher = base64.b64decode(encoded.encode("ascii"))
        _, plain = win32crypt.CryptUnprotectData(
            cipher, None, None, None, _CRYPTPROTECT_UI_FORBIDDEN
        )
    except Exception as e:
        raise RuntimeError(f"DPAPI 解密失败（可能已换账户/机器或重装系统）: {e}") from e
    return plain.decode("utf-8")


def load_api_key() -> str:
    """读取 API Key。优先解密 api_key_encrypted，失败视为未配置；兼容迁移旧明文。"""
    cfg = load_config()

    encrypted = (cfg.get("api_key_encrypted") or "").strip()
    if encrypted:
        try:
            return _decrypt(encrypted).strip()
        except Exception:
            # 换机/换账户/重装后密文无法解密，视为未配置，返回空（设置页引导重输）
            return ""

    plaintext = (cfg.get("api_key") or "").strip()
    if plaintext:
        # 旧版明文自动迁移：加密重写并清除明文
        try:
            save_api_key(plaintext)
        except Exception:
            # 迁移失败（如非 Windows 环境）时保留明文可用，避免旧数据丢失
            pass
        return plaintext
    return ""


def save_api_key(key: str) -> None:
    """加密保存 API Key。空值等同清除。"""
    key = (key or "").strip()
    if not key:
        clear_api_key()
        return
    encrypted = _encrypt(key)
    cfg = load_config()
    cfg["api_key_encrypted"] = encrypted
    cfg["api_key"] = ""
    save_config(cfg)


def clear_api_key() -> None:
    """清除已保存的 API Key（置空，保留配置文件结构）。"""
    cfg = load_config()
    cfg["api_key"] = ""
    cfg["api_key_encrypted"] = ""
    save_config(cfg)


def has_api_key() -> bool:
    return bool(load_api_key())


def load_base_url() -> str:
    """读取自定义模型服务地址（非敏感，明文保存）。为空表示使用默认地址。"""
    return (load_config().get("base_url") or "").strip()


def save_base_url(url: str) -> None:
    cfg = load_config()
    cfg["base_url"] = (url or "").strip()
    save_config(cfg)


def get_output_root() -> Path:
    """PPT 输出根目录：固定锚定到项目目录下的 debug_logs 文件夹（绝对定位，不受启动 cwd 影响）。"""
    out = Path(__file__).resolve().parent / "debug_logs"
    out.mkdir(parents=True, exist_ok=True)
    return out
