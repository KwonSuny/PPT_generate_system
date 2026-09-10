# PPT 自动生成系统

基于多智能体协作的 PowerPoint 演示文稿生成系统。系统调用大语言模型完成内容规划、文案撰写、配图生成与视觉排版，最终通过 python-pptx 渲染为 PPTX 文件，并提供桌面窗口与 Web 两种交互入口。

## 核心特性

- 多智能体流水线：Planner（大纲规划）、Writer（文案扩写）、Reviewer（质检）、ImageGenerator（配图生成）、Layout（视觉排版）、Renderer（物理渲染）。
- 页面级增量修改：支持指定单页、多页或全部页进行修改，其余页面沿用缓存，避免全量重生成。
- 字段级锁定：区分"仅修改文本"与"仅调整设计"，从软提示升级为清洗器层面的硬约束。
- 逐页排版：Layout 按页生成物理蓝图，规避单次大请求触发网关断连。
- 会话状态归档：每次生成对应一个独立的时间戳文件夹，新建对话不携带历史信息。
- 离线演示模式：未配置 API Key 时生成基础 PPT，便于本地调试与演示。

## 项目结构

```
PPT_system\
├── main.py                    # 生成管线主控（全量 + 增量）
├── app.py                     # Flask Web 应用与会话管理
├── renderer.py                # PPTX 物理渲染引擎
├── launcher.py                # 桌面启动器（pywebview 原生窗口）
├── config_manager.py          # API Key 持久化与输出目录定位
├── agents/
│   ├── planner.py             # 大纲规划智能体
│   ├── writer.py              # 文案扩写智能体
│   ├── review.py              # 质检审查（页数、SVG、文本长度校验）
│   ├── image_generator.py     # 配图生成智能体
│   └── template.py            # 视觉排版智能体（Layout）
├── config/
│   ├── schema.py              # Pydantic 数据模型
│   ├── state.py               # 会话状态管理
│   ├── sanitizers.py          # 数据清洗与字段级锁定
│   ├── debug.py               # 调试日志工具
│   ├── llm_client.py          # LLM 客户端封装（超时 + 重试）
│   └── config.json            # API Key 配置（运行时生成）
├── templates/
│   ├── index.html             # 主界面
│   └── setup.html             # API Key 设置页
├── test_incremental.py        # 增量管线测试
├── test_full_layout_paging.py # 全量逐页排版测试
├── test_app_routing.py        # Flask 路由测试
├── test_desktop.py            # 桌面封装测试
├── debug_logs/                # 运行输出目录（见"输出与日志"）
```

## 环境要求与依赖

- Python 3.9 及以上（开发环境为 3.12）。
- 阿里云百炼（DashScope）MaaS 兼容接口的 API Key，用于调用 Qwen 系列文本模型与多模态生图模型。

第三方依赖：

```bash
pip install flask python-pptx pydantic openai dashscope requests cairosvg pywebview pywin32
```

各依赖用途：

| 依赖 | 用途 |
| --- | --- |
| flask | Web 服务与路由 |
| python-pptx | PPTX 物理渲染 |
| pydantic | 强类型数据模型与校验 |
| openai | 以 OpenAI 兼容协议调用文本模型 |
| dashscope | 多模态生图模型调用 |
| requests | 下载生成的图片 |
| cairosvg | SVG 转 PNG |
| pywebview | 桌面原生窗口封装（仅 launcher.py 需要） |

系统级依赖：

- cairosvg 底层依赖 cairo 动态库。Windows 下使用 conda 环境时，cairo DLL 位于环境安装目录的 `Library\bin` 下，运行前需将该目录加入 PATH（激活 conda 环境即已包含）。缺少 cairo 库时会报 `no library called "cairo-2"`。

## 配置 API Key

生成管线统一从环境变量 `DASHSCOPE_API_KEY` 读取 Key，该变量由入口程序在启动时注入：

- 桌面启动器（launcher.py）：从 `config/config.json` 读取已保存的 Key 并写入进程环境变量；未配置时进入设置页。
- 设置页保存：调用 `/api/save_key`，将 Key 写入 `config/config.json` 并同步到进程环境变量。
- 命令行直接运行（main.py / app.py）：可手动设置环境变量 `DASHSCOPE_API_KEY`。

Key 在 `config/config.json` 中以加密形式存储：使用 Windows 数据保护接口（DPAPI，通过 pywin32 的 win32crypt 调用），密文绑定当前 Windows 用户与机器。换账户、换机器或重装系统后密文无法解密，系统会视为未配置并引导重新输入一次 Key。设置页与接口均不回显完整 Key。

主界面提供"注销 API Key"入口，调用 `/api/logout` 清除本机密钥并返回设置页，便于更换 Key。

## 自定义模型服务地址

默认服务地址为阿里云百炼 MaaS 兼容接口（定义于 `config/llm_client.py` 的 `DEFAULT_BASE_URL`）。如需接入个人或自建的 OpenAI 兼容接口，可在设置页填写"模型服务地址"：

- 填写后与 Key 一起保存到 `config/config.json` 的 `base_url` 字段（明文存储，非敏感），并写入进程环境变量 `DASHSCOPE_BASE_URL`。
- 留空则使用默认地址。
- 也可直接设置环境变量 `DASHSCOPE_BASE_URL` 手动指定。
- 三个智能体（Planner / Writer / Layout）创建客户端时按「显式传参 > 环境变量 > 默认地址」的优先级取用。

## 使用说明

### 方式一：桌面窗口（推荐）

```bash
python launcher.py
```

启动后弹出独立原生窗口。首次运行进入设置页填写 API Key，之后直接进入主界面。

### 方式二：Web 界面

```bash
python app.py
```

浏览器访问 `http://127.0.0.1:8000`。服务仅监听本机回环地址，避免局域网暴露配置接口。

### 方式三：命令行直接生成

编辑 `main.py` 底部的调用，指定主题与输出文件名后运行：

```bash
python main.py
```

## 系统架构

### 智能体流水线

全量生成按以下顺序执行：

1. Planner：根据主题生成 PPT 大纲，规划每页布局类型与配色。
2. Writer：将大纲扩写为详细文案，匹配图标，生成 SVG 代码与配图描述。
3. Reviewer：校验页数（1-5 页）、SVG 合规性、单条文本长度，最多 3 轮质检闭环。
4. ImageGenerator：按页生成配图；含 SVG 的页面跳过本地生图。
5. Layout：逐页生成物理画布蓝图，计算图元坐标、尺寸与颜色。
6. Renderer：使用 python-pptx 将蓝图渲染为 PPTX 文件。

### 数据模型

核心 Pydantic 模型定义于 `config/schema.py`：

- `ColorPalette`：单页配色方案。
- `SlideElement`：单页幻灯片内容。
- `PPTOutline`：整份 PPT 的语义大纲。
- `CanvasElement`：物理画布元素。
- `SlideCanvas`：单页物理画布。
- `PPTFinalBlueprint`：最终渲染蓝图。

## 生成模式与增量修改

Web 界面支持四种模式：

| 模式 | 行为 |
| --- | --- |
| new | 全新生成，不继承任何历史信息 |
| continue | 基于历史继续修改（不指定页码时走全量管线） |
| modify_text | 仅修改文本，锁定版式与配色 |
| modify_design | 仅调整设计与配色，锁定文案 |

页码输入支持单页、多页、范围与混合写法，例如 `2`、`2,3`、`2-4`、`1,3-5`。指定页码时走增量管线，仅对目标页调用 Writer / 生图 / Layout，其余页沿用缓存；`modify_text` / `modify_design` 不填页码时执行全局修改（全部页，跳过 Planner）。

文本与设计的区分由 `config/sanitizers.py` 的字段级锁定实现，不依赖提示词软约束：

- text 锁：锁定 `layout_type`、`theme`，允许修改标题、要点、正文等文本字段。
- design 锁：锁定标题、摘要、要点、正文，允许修改 `theme` 与 `layout_type`。

## 缓存机制

`run_ppt_system` 提供两个缓存注入参数，用于调试时跳过部分阶段：

- `from_content_json`：注入大纲 JSON，跳过 Planner / Writer / 生图，直接进入排版。
- `from_blueprint_json`：注入蓝图 JSON，跳过全部生成阶段，直接渲染。

## 输出与日志

所有生成物统一收纳在项目目录下的 `debug_logs` 文件夹（可通过 `config_manager.get_output_root` 调整），每次生成对应一个独立的大时间戳文件夹：

```
debug_logs/
├── current_session.txt          # 会话指针，记录当前 run_id
└── <YYYYMMDD_HHMMSS>/           # 单次生成的大时间戳文件夹
    ├── <name>.pptx              # 生成的 PPT
    ├── state.json               # 本次对话状态（随对话归档）
    ├── generated/               # 应用运行态
    │   ├── run_<name>.log
    │   └── snapshots/<name>/
    │       ├── outline.json     # 语义大纲快照
    │       └── blueprint.json   # 物理蓝图快照
    └── run/                     # Agent 调试日志
        ├── raw_outline_attempt_*.json
        ├── compiled_content_attempt_*.json
        ├── final_content_*.json
        ├── final_blueprint.json
        ├── slide_accent_*.png
        └── error.txt / error.json
```

`state.json` 随每次对话归档到对应的时间戳文件夹内，代表单次对话。点击"新建对话"会清除会话指针，下一次加载不再携带上一轮的历史信息。快照文件（`outline.json`、`blueprint.json`）用于支持后续增量修改的回放。

## Web API 接口

- `GET /`：主界面。
- `GET /setup`：API Key 设置页。
- `GET /api/has_key`：查询是否已配置 Key。
- `POST /api/save_key`：保存 Key。
- `POST /api/logout`：注销 Key（清除本机密钥并返回设置页）。
- `POST /generate`：生成 PPT。
- `GET /state`：获取生成状态与进度。
- `POST /api/reset`：重置会话。

## 测试

项目包含 4 个测试脚本，均不依赖真实 LLM API：

```bash
python test_incremental.py          # 增量管线逻辑
python test_full_layout_paging.py   # 全量逐页排版逻辑
python test_app_routing.py          # Flask 路由与会话归档
python test_desktop.py              # 桌面封装与设置页
```

Windows 下若 cairosvg 依赖 cairo DLL，运行测试前需将 conda 环境的 `Library\bin` 加入 PATH。

## 注意事项与已知限制

1. 未配置 API Key 时，系统自动进入离线演示模式，生成基础 PPT。
2. 页数限制为 1-5 页，校验逻辑位于 `agents/review.py`。
3. 逐页排版使 Layout 调用次数与页数成正比，总耗时随页数线性增长，但单次请求体积小、稳定性高。
4. 中文字体在 SVG 渲染中需使用支持中文的字体族。
5. `config/config.json` 与 `debug_logs/` 中的 `current_session.txt` 为运行期生成文件，可随时删除；前者清除后需重新配置 Key，后者清除后会话历史清空。
