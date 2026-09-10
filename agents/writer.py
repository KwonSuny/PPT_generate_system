import json
from config.schema import PPTOutline
from config.sanitizers import normalize_outline_payload, parse_json_object
from config.state import PPTGenerationState
from config.llm_client import create_client, chat_completion_with_retry


def get_writer_agent(api_key: str, model_name: str = "qwen3.7-max"):
    """
    创建并返回一个负责长文本扩写、精准图标指派与高级视觉元素补全的文案打磨智能体。
    """
    client = create_client(api_key)
    
    def write_ppt_content(state: PPTGenerationState, planner_output_json: dict) -> PPTOutline:
        print(" [Writer Agent] 收到大纲骨架，正在深入撰写长文本、分配微图标并规划视觉元素...")
        
        # 将 planner 的标准输出转化为字符串输入给模型
        if isinstance(planner_output_json, dict):
            input_str = json.dumps(planner_output_json, ensure_ascii=False, indent=2)
        else:
            input_str = str(planner_output_json)
            
        system_prompt = """
        # ROLE
        你是一个精通内容叙事与数字化视觉编排的顶级演示文稿架构师。
        你的任务是：接棒 Planner 生成的高层大纲，将其深度扩写，并完美补全所有视觉与文本细节。

        # TASK
        输入是一个包含基本大纲和排版类型的 JSON 对象。你必须遍历其中的 `slides` 数组，保持原有的页面顺序、数量和基础字段（如 title, punchline, layout_type, theme）绝对不变，并【精准补全】以下核心高价值字段：
        1. `full_content`: 根据 `bullet_points` 的大纲要点，扩写出详细的、极具宣讲说服力的文本段落。数量必须与 `bullet_points` 严格一一对应，字数控制在 120 字以内，严禁照抄 'bullet_points' 的原文。
        2. `icons`: 为当前页面的每个 bullet 要点匹配一个最具现代化语境的 Lucide 图标库名称（例如：科技/算力用 'cpu'，安全/防御用 'shield'，数据/趋势用 'trending-up'，用户/团队用 'user'，全球化用 'globe'，高能/破局用 'rocket'）。图标数组的长度必须与 `bullet_points` 数量严格完全一致！
        3. `svg_code`: 当且仅当 `layout_type == 'svg'` 时，在此处填入精美的、具备现代极简 UI/UX 设计感的矢量图形代码（包含高级流程图、系统架构图或结构化高颜值数据表格）。其他布局请设为 null。
        4. `image_prompt`: 为当前页面提炼出一段适合生成现代高级演示文稿配图的中文视觉意境描述词。如果没有配图需求，请设为 null。
        5. `image_path`: 统一保持为 null，由后面的物理生图智能体自动填充。

        #  OUTPUT RULES (绝对合规红线) 
        1. 必须输出一个且仅一个严格合法的标准 JSON 对象，其根结构必须是 `{"topic": "...", "slides": [...]}`。
        2. `slides` 数组中的每一个元素必须是且只能是一个包含完整字段的 JSON 对象（{}），代表一页幻灯片。
        3. 【绝对禁止】在 `slides` 数组中直接输出任何独立的字符串（如属性字段名 'image_path', 'icons', 'title' 等）！`slides` 数组里只能存在幻灯片对象，严禁夹带任何 schema 声明片段作为独立的数组元素！
        4. 确保 JSON 语法绝对正确，不要有多余的逗号，不要提前闭合括号。

        # SVG MODERN AESTHETIC RULES 
        当绘制 `svg_code` 时，必须彻底丢弃传统的“工科死板方块连线”，全面实施以下现代化互联网高端 UI 设计规范：
        1. **渐变色填充 (Gradients)**: 严禁使用任何纯单色填充！必须在 `<defs>` 区域内定义 `<linearGradient>` 或 `<radialGradient>`。例如设置多色锚点过渡（如从深邃色到亮炫色），并在形状的 `fill` 属性中使用 `url(#gradient_id)` 引用。
        2. **柔和阴影与立体滤镜 (Drop Shadows & Blur)**: 必须为流程图的关键卡片、核心节点加上柔和的高斯模糊阴影滤镜，例如使用 `<filter id="soft-shadow">` 结合 `<feDropShadow dx="0" dy="4" stdDeviation="6" flood-opacity="0.08" flood-color="#000000"/>`，营造完美的悬浮层级感（Elevation）。
        3. **顺滑曲线与圆角流 (Smooth Curves & Roundness)**: 所有的矩形卡片必须配合非常明显的平滑大圆角（`rx="16" ry="16"` 或 `rx="20" ry="20"`）。各节点之间的连接线严禁使用生硬的绝对直线或 90 度死板折线，必须使用三次贝塞尔曲线路径（如 `<path d="M... C..."/>`），使其具有水流般、生动的流线型导向。
        4. **高级高颜值表格与对比矩阵规范 (Advanced SVG Tables)**:
           - 若当前页面属于表格数据或指标对比维度场景，必须使用原生 SVG 标签精细构建完整的网格与文本。
           - **表头美化**: 表头区域使用深色或显明的品牌主色调（如 `theme.primary`）作为背景矩形填充，文字显式加粗，颜色设为高对比浅色（如 `#FFFFFF`）。
           - **数据行与斑马纹**: 数据行的背景必须使用交替轻微色区分（如纯白 `#FFFFFF` 与极其柔和的浅灰/浅蓝 `#F8FAFC` 互相交替）以建立杂志级透气感。行与行、列与列之间使用极细的浅色分割线（如 `stroke="#E2E8F0"`，`stroke-width="1"`）。
           - **坐标微调与防重叠**: 每个单元格内的 `<text>` 标签必须精确分配唯一的 `x` 和 `y` 坐标，确保文本垂直水平完美居中或优雅左对齐。必须留足单元格内边距，绝对不能发生任何文字重叠、出格或中途被截断。
        5. SVG 内所有文字必须留足内边距，不得贴边、溢出卡片或与连线/节点重叠。
        # 结构完整性铁律（防止 Pydantic 校验失败的核心规则）
        structural integrity rules 
        你输出的 JSON 必须完美继承输入的所有骨架数据。对于 `slides` 数组中的每一页幻灯片，你必须：
        1. **绝对原封不动地保留并复制** 输入中的 `title`、`punchline`、`bullet_points`、`layout_type` 和 `theme` 字段。严禁删除或遗漏它们！
        2. 仅对原先为空的 `full_content`、`svg_code`、`image_prompt`、`image_path` 进行生成、扩写或保持 null。

        # output format
        1. 严禁在 SVG 内的 `<text>` 标签中使用 font-family="Arial" 或 font-family="sans-serif" 等不支持中文的纯英文字体！
        2. 所有 `<text>` 标签必须统一显式指定支持中文的现代系统字体族：
        font-family="'Microsoft YaHei', 'SimHei', 'PingFang SC'"
        例如：`<text x="100" y="50" font-family="'Microsoft YaHei', 'SimHei', 'PingFang SC'" font-size="16" fill="#FFFFFF">核心原理解析</text>`
        3. 确保文字的 x, y 坐标与字号（font-size）搭配合理，防止高分辨率渲染时由于中文字体较宽发生文字重叠。
        4. `svg_code` 必须是完整、可被 XML 解析器解析的纯 SVG 字符串：必须包含 `<svg xmlns="[http://www.w3.org/2000/svg](http://www.w3.org/2000/svg)" viewBox="0 0 1280 720">`，禁止 Markdown 代码块围栏。
        5. SVG 文本中的 XML 特殊字符必须转义：`&` 写成 `&amp;`，`<` 写成 `&lt;`，`>` 写成 `&gt;`。禁止出现未转义的裸 `&`。
        6. 禁止使用 `<foreignObject>`、外链图片、外链字体、脚本、动画和任何需要联网加载的资源；所有视觉必须由 SVG 原生图形、路径、渐变、滤镜和文本完成。
        
        """

        mode_hint = ""
        if state.operation_mode == "incremental_page":
            pages_desc = state.page_label()
            if state.modify_scope == "design":
                mode_hint = (
                    f"本次为针对 {pages_desc} 的单页「仅设计」修改：只允许调整该页的设计风格——"
                    "theme 的 primary/secondary/content/background 四个色值，以及 layout_type 版式。"
                    "配色可参考输入中的 neighbor_themes 保持同一色彩家族连续演进。"
                    "严禁修改 title、punchline、bullet_points、full_content 等任何文本字段。"
                    "输出必须仍然是仅含 1 个元素的 slides 数组。"
                )
            else:
                mode_hint = (
                    f"本次为针对 {pages_desc} 的单页增量修改：输入 JSON 只包含该页内容。"
                    "请依据用户最新要求更新该页的 title、punchline、bullet_points、full_content、image_prompt 等文本字段，"
                    "theme 需与输入中的 neighbor_themes（邻页配色）保持同一色彩家族的连续演进。"
                    "输出必须仍然是仅含 1 个元素的 slides 数组，严禁增删页面或输出多页。"
                )
        elif state.operation_mode == "continue":
            mode_hint = "你将基于现有 PPT 内容继续优化，不要全量重写。"
        elif state.operation_mode == "modify_text":
            mode_hint = "请仅修改文本内容，保留原有版式与布局结构。"
        elif state.operation_mode == "modify_design":
            mode_hint = "请仅调整设计风格与视觉细节，不更改核心文案。"
        if (
            state.modify_page is not None
            and state.operation_mode not in ("incremental_page",)
        ):
            mode_hint += f" 重点聚焦第 {state.modify_page} 页。"

        state.append_progress("Writer: 补全文案与视觉细节")
        response = chat_completion_with_retry(
            client,
            model=model_name,
            messages=state.build_messages(
                system_prompt,
                f"{mode_hint}\n\n这是 Planner 规划出的初始大纲，请按照提示词补充相应字段，严格确保 slides 数组中只包含完整的幻灯片对象：\n\n{input_str}"
            ),
            response_format={"type": "json_object"},
            temperature=0.3,
            extra_body={"enable_search": True, "enable_thinking": True},
        )

        if not hasattr(response, 'choices') or response.choices is None or len(response.choices) == 0:
            print("\n [Writer Agent 崩溃警报] DashScope API 没有返回任何有效文本！")
            raise ValueError("DashScope 接口返回了空 choices。")
        
        compiled_json = response.choices[0].message.content.strip()
        state.append_message("assistant", compiled_json)
        
        try:
            if compiled_json.startswith("```json"):
                compiled_json = compiled_json.split("```json")[1].split("```")[0].strip()
            elif compiled_json.startswith("```"):
                compiled_json = compiled_json.split("```")[1].split("```")[0].strip()
            
            data = parse_json_object(compiled_json)
            
            if "slides" in data and isinstance(data["slides"], list):
                clean_slides = [slide for slide in data["slides"] if isinstance(slide, dict)]
                if len(clean_slides) != len(data["slides"]):
                    data["slides"] = clean_slides 
            
            # 根据操作模式与修改范围决定字段级锁定策略：
            #   - 增量（单页/多页/全局）：design 锁文字、text 锁版式、both 全解锁
            #   - 全量（new/continue/回退）：锁骨架（structure），Writer 只扩写正文
            if state.operation_mode == "incremental_page":
                if state.modify_scope == "design":
                    lock_scope = "design"
                elif state.modify_scope == "text":
                    lock_scope = "text"
                else:
                    lock_scope = None
            else:
                lock_scope = "structure"
            normalized = normalize_outline_payload(
                data,
                skeleton=planner_output_json,
                lock_scope=lock_scope,
            )
            outline = PPTOutline.model_validate(normalized)
            state.set("outline", outline.model_dump())
            return outline
            
        except Exception as e:
            print(f"[Writer Agent] 结构自动清洗过程发生异常: {e}，将使用 Planner 骨架做保底回填...")
            normalized = normalize_outline_payload(
                {},
                skeleton=planner_output_json,
                lock_scope="structure",
            )
            outline = PPTOutline.model_validate(normalized)
            state.set("outline", outline.model_dump())
            return outline

    return write_ppt_content