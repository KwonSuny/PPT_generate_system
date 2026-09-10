# -*- coding: utf-8 -*-
"""桌面化改造的安全/逻辑测试：验证设置页、Key 保存接口、输出目录迁移与增量路由不回归。"""
import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ.pop("DASHSCOPE_API_KEY", None)
os.environ.pop("DASHSCOPE_BASE_URL", None)

# 隔离：把配置与输出目录重定向到临时目录，避免污染用户真实数据
import config_manager
import app as app_module

_tmp = Path(tempfile.mkdtemp(prefix="ppt_desktop_test_"))
config_manager.CONFIG_DIR = _tmp / "cfg"
config_manager.CONFIG_FILE = config_manager.CONFIG_DIR / "config.json"
app_module.OUTPUT_ROOT = _tmp / "output"
app_module.OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

client = app_module.app.test_client()


def check(name, cond):
    print(("PASS: " if cond else "FAIL: ") + name)
    if not cond:
        raise SystemExit(1)


# 1) 设置页可达
r = client.get("/setup")
check("GET /setup 返回 200", r.status_code == 200)
body = r.get_data(as_text=True)
check("设置页含品牌标题", "PPT 生成器" in body and "DashScope API Key" in body)

# 2) 初始无 key
r = client.get("/api/has_key")
check("初始 has_key=False", r.status_code == 200 and r.get_json()["has_key"] is False)

# 3) 空 key 提交应被拒绝
r = client.post("/api/save_key", json={"api_key": "   "})
check("空 key 返回 400", r.status_code == 400)

# 4) 有效 key 保存（含自定义服务地址）
TEST_KEY = "sk-test-1234567890"
TEST_BASE_URL = "https://my-personal-endpoint.example.com/v1"
r = client.post("/api/save_key", json={"api_key": TEST_KEY, "base_url": TEST_BASE_URL})
check("保存 key 返回 success", r.status_code == 200 and r.get_json()["status"] == "success")
check("key 写入配置文件", config_manager.load_api_key() == TEST_KEY)
check("key 同步到环境变量", os.environ.get("DASHSCOPE_API_KEY") == TEST_KEY)
check("base_url 写入配置文件", config_manager.load_base_url() == TEST_BASE_URL)
check("base_url 同步到环境变量", os.environ.get("DASHSCOPE_BASE_URL") == TEST_BASE_URL)
raw_cfg = config_manager.CONFIG_FILE.read_text(encoding="utf-8")
check("config.json 不含明文 key", TEST_KEY not in raw_cfg)
check("config.json 含加密字段", "api_key_encrypted" in raw_cfg)
r = client.get("/api/has_key")
check("保存后 has_key=True", r.get_json()["has_key"] is True)
check("has_key 返回 base_url 供预填", r.get_json().get("base_url") == TEST_BASE_URL)

# 5) key 不会回显
r = client.get("/api/has_key")
check("接口不回显完整 key", TEST_KEY not in r.get_data(as_text=True))

# 6) 输出目录（用户已改为项目目录下的 debug_logs）
root = config_manager.get_output_root()
check("输出根目录为 debug_logs", root.name == "debug_logs")
check("输出根目录已创建", root.exists())

# 7) 首页与生成路由未回归（预置缓存大纲后走增量路由）
outline = {
    "topic": "测试",
    "slides": [
        {"title": "第1页", "punchline": "一", "bullet_points": ["a"], "full_content": ["a"],
         "layout_type": "standard_text", "theme": {"primary": "#1E3A8A", "secondary": "#10B981",
         "content": "#64748B", "background": "#F8FAFC"}, "svg_code": None, "image_prompt": None,
         "image_path": None, "icons": []},
        {"title": "第2页", "punchline": "二", "bullet_points": ["b"], "full_content": ["b"],
         "layout_type": "standard_text", "theme": {"primary": "#1E3A8A", "secondary": "#10B981",
         "content": "#64748B", "background": "#F8FAFC"}, "svg_code": None, "image_prompt": None,
         "image_path": None, "icons": []},
    ],
}
from config.state import PPTGenerationState
seed = PPTGenerationState(topic="测试", requirements="初始")
seed.set("outline", outline)
app_module._save_state(seed)

r = client.post("/generate", data={
    "topic": "测试", "requirements": "改第2页", "filename": "demo.pptx",
    "mode": "modify_text", "modify_page": "2",
})
check("POST /generate 返回 200", r.status_code == 200)
check("页面显示单页仅改文本标签", "单页仅改文本" in r.get_data(as_text=True))

print("\nALL DESKTOP-WRAP TESTS PASSED")
print(f"临时隔离目录: {_tmp}")
