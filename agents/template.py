import json
from config.schema import PPTFinalBlueprint
from config.sanitizers import normalize_blueprint_payload, parse_json_object
from config.state import PPTGenerationState
from config.llm_client import create_client, chat_completion_with_retry


def get_layout_agent(api_key: str, model_name: str = "qwen3.8-max"):
    """
    创建并返回一个精通数字化网格排版与高级原件堆叠的画布物理精算智能体。
    """
    client = create_client(api_key)
    
    def design_layout(state: PPTGenerationState, writer_output_json: dict) -> PPTFinalBlueprint:
        print(" [Layout Agent] 感知单页配色、微图标与图片路径，开始进行排版")
        
        if isinstance(writer_output_json, dict):
            input_str = json.dumps(writer_output_json, ensure_ascii=False, indent=2)
        else:
            input_str = str(writer_output_json)

        system_prompt = """
        # ROLE
        你是一个精通数字化网格排版系统（Grid System）和高端幻灯片版式的顶尖视觉总监。
        你的任务是：将输入的包含独立配色、专属图标序列以及本地点缀图片路径的“语义大纲”彻底转化为完全确定且多样布局的“物理画布蓝图（JSON 对象）”。

        # CANVAS SPECIFICATION (画布物理参数)
        - 画布大小固定为标准的 16:9 宽屏：宽度 = 13.333 英寸，高度 = 7.5 英寸。
        - 所有的物理图元坐标 (`left`, `top`, `width`, `height`) 必须基于此范围精算，单位为浮点数（英寸）。

        #  STRICT SCHEMA LIMITATIONS (绝对不可逾越的类型红线) 
        1. 你的物理蓝图数据必须 100% 完美对齐 Pydantic 校验器，严禁自行发明或简化任何字段值！
        2. 必须保留原有的数据(如图片、表格等)，除非无法排版。

        # OUTPUT FORMAT SPECIFICATION (严格输出规范)
        1. 输出必须是一个标准的 JSON 对象，其根节点必须包含且仅包含一个 `"canvases"` 键，其值为主体画布的数组（即：{"canvases": [...] }）。
        2. 严禁将整个响应或外部结构伪装、包裹成任何形如 `{"type": "textbox"}` 的多余图元节点。你输出的是整份幻灯片的总物理蓝图。
        3. 检查所有元素是否有遮盖，必要时进行放缩、移动或重新布局，确保所有图元都在画布范围内且不重叠。

        ## 1. 【图元类型 type 绝对限制】
        `CanvasElement` 中的 `type` 字段只能且必须是以下五个基础字符串之一：['textbox', 'shape', 'svg', 'image', 'icon']。
        -  严禁输出 `"type": "text"`！所有的文本框、标题、正文段落必须声明为 `"type": "textbox"`。
        -  **新增 `"type": "icon"`**：用于高档微图标点缀。当检测到输入数据中含有该页面的 `icons` 数组时，你必须在对应文本内容的左侧或上方，合理规划出一个 `type: "icon"` 的物理图元，填入对应的 `icon_name` 字符串（如 'shield', 'cpu'）并指定 `icon_color`。

        ## 2. 【形状类型 shape_type 绝对限制】
        当且仅当 `type` 为 `"shape"` 时，`shape_type` 只能且必须是以下五个基础字符串之一：['rectangle', 'rounded_rectangle', 'ellipse', 'line', 'arrow']。
        - 严禁输出 `"shape_type": "circle"`！若需绘制正圆，声明为 `"shape_type": "ellipse"` 且确保 `width` 和 `height`完全相等。
        -  **新增 `"shape_type": "line"`**：用于在页面大标题下方、或者多栏内容之间绘制一条精细的视觉装饰分割线（如宽 10.0，高 0.04），颜色采用 `theme.secondary`，使页面极极杂志呼吸感。
        -  **新增 `"shape_type": "arrow"`**：用于步骤递进、历史演进或时间轴（`layout_type == 'timeline'`）的场景，充当各个内容卡片之间的物理流程流向引导箭头。箭头不能太大，防止遮挡。


        #  VISUAL ELEMENT UPGRADE RULES (高级视觉元组件拼装规则)
        拒绝单调粗暴的纯文字排版，你要像玩乐高一样，利用新图元进行“复合组件”的物理堆叠：
        1. **现代高端容器卡片堆叠法**：
           当排版多栏内容（`column`）或时间轴（`timeline`）时，不要干巴巴地放字。应当先画一个浅色或带有边框的圆角矩形作为容器底衬 (`type="shape", shape_type="rounded_rectangle", fill_color=微亮背景色`)，然后在这块物理区域内部，精准重叠计算并摆放：一个 `type="icon"` 图标组件 + 一个带有精美标题和正文的 `type="textbox"`。
        2. **图标与文本框的物理网格对齐（对齐铁律）**：
           - **左右并排对齐**：若图标放在文字左侧，紧挨着它的文本框必须预留至少 0.2 英寸呼吸间距，且文本框的 `top` 应与图标物理居中对齐。
           - **上下堆叠对齐**：若图标放在文字上方，则下方的文本框必须预留至少 0.6 英寸呼吸间距。
        3. **SVG / 复杂表格图元空间精算铁律**：
           - 当检测到该幻灯片包含已生成的 `svg_code`（如表格或高级拓扑图）时，必须在画布中规划一个核心的 `type: "svg"` 物理图元。
           - 必须为该表格/图形组件开辟出极其宽敞、显眼的独立黄金视区，使其成为当前全页的视觉主体。
           - 必须精准计算它与上方的页面小标题框（`textbox`）的位置边界，确保留有至少 0.2 英寸的绝对空气空隙，彻底杜绝任何层级重叠。
           - 必须按比例缩放 SVG 。
        # LAYOUT RULES (排版精算规则)
        - 保持原有的多栏算框、正文页2/3安全网格“侧边点缀排版法”等业务逻辑。
        - **单页配色 (`theme`)**: 每一个 slide 都带有一个独一无二的 `theme`。你必须将 `theme.background` 作为该页 `SlideCanvas` 的 `background_color`，并将相应的颜色分发给各个文本框的 `font_color`、图标的 `icon_color` 或形状的 `fill_color`、`line_color`。
        - 所有图元必须严格落在 13.333 x 7.5 英寸画布内，禁止负坐标、禁止超出右边界或下边界，元素之间至少保留 0.1 英寸安全距离。
        - `image` 只代表美观插图/点缀图，不是背景图。严禁将 image 铺满全屏或作为文字衬底。
        - image 必须使用侧边栏、角落或独立视觉区；不要穿过正文列或覆盖标题。

        # READABILITY AND IMAGE RULES (可读性与图片比例规则)
        - 文本与背景的对比度必须足够高：浅色背景使用 #111827/#334155 等深色文字；深色背景使用 #FFFFFF/#E5E7EB 等浅色文字。
        - 标题框高度不得低于 0.65 英寸，正文段落框高度不得低于 0.45 英寸；多行正文必须预留上下内边距，不要让文字贴边。
        - 任何 image 图元都必须保持原始宽高比，按 contain 思路放入插图区。

        # TEXTBOX PARAGRAPH RULES (文本框段落规范)
        1. 严禁将任何 type == "textbox" 的元素的 "paragraphs" 字段设为 null！
        2. 你在对当前页面进行排版布控算框时，必须把输入数据中当前页面的 `title`、`punchline` 或 `full_content` 中的每一条文本，原汁原味地转换成符合 `TextParagraph` Schema 的结构，填入对应 textbox 的 paragraphs 列表中！
        """

        mode_hint = ""
        if state.operation_mode == "incremental_page":
            pages_desc = state.page_label()
            mode_hint = (
                f"本次为针对 {pages_desc} 的单页增量排版：输入 JSON 只包含目标页内容。"
                "你输出的 canvases 数组长度必须严格等于 1（仅该页的物理画布），严禁输出多页。"
                "配色需与输入中的 neighbor_themes（邻页配色）保持同一色彩家族的连续演进。"
            )
        elif state.operation_mode == "continue":
            mode_hint = "你将基于现有 PPT 继续优化视觉布局，不要完全重置页面结构。"
        elif state.operation_mode == "modify_text":
            mode_hint = "本轮仅调整文本内容，视觉排版尽量保持当前布局。"
        elif state.operation_mode == "modify_design":
            mode_hint = "本轮仅优化视觉风格与排版细节，不要更改核心文本。"
        if state.modify_page is not None and state.operation_mode not in ("incremental_page",):
            mode_hint += f" 重点聚焦第 {state.modify_page} 页。"

        # 单页输入约束：无论何种 operation_mode（含全量管线逐页排版时 mode 仍为 new/continue），
        # 只要本次输入仅含 1 页内容，就必须要求输出严格 1 个画布，防止模型对单页输入误输出多页。
        input_slides = (writer_output_json or {}).get("slides") if isinstance(writer_output_json, dict) else None
        if isinstance(input_slides, list) and len(input_slides) == 1:
            single_hint = (
                "本次输入仅包含 1 页内容，你输出的 canvases 数组长度必须严格等于 1"
                "（仅该页的物理画布），严禁输出多页。"
            )
            mode_hint = f"{mode_hint} {single_hint}".strip()

        state.append_progress("Layout: 生成物理蓝图")
        # 精简上下文：Layout 只需语义大纲本身即可排版。避免把完整对话历史与
        # current_outline/current_blueprint 重复塞入（payload 过大是此前触发
        # 「Server disconnected without sending a response」的关键诱因）。
        lean_messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": (
                    f"{mode_hint}\n\n请阅读以下语义大纲与主题色彩，并为每一页精算物理图元布局。"
                    f"充分利用新增加的 icon、line 和 arrow 元素来丰富页面质感：\n\n{input_str}"
                ),
            },
        ]
        response = chat_completion_with_retry(
            client,
            model=model_name,
            messages=lean_messages,
            response_format={"type": "json_object"},
            temperature=0.1,
        )
        
        blueprint_json = response.choices[0].message.content.strip()
        state.append_message("assistant", blueprint_json)

        try:
            cleaned_json = blueprint_json.strip()
            if cleaned_json.startswith("```json"):
                cleaned_json = cleaned_json.split("```json")[1].split("```")[0].strip()
            elif cleaned_json.startswith("```"):
                cleaned_json = cleaned_json.split("```")[1].split("```")[0].strip()
                
            parsed_data = parse_json_object(cleaned_json)
            
            if isinstance(parsed_data, dict) and "type" in parsed_data and "text" in parsed_data:
                inner_content = parsed_data["text"].strip()
                parsed_data = json.loads(inner_content)
                
            if isinstance(parsed_data, list):
                parsed_data = {"canvases": parsed_data}
                
            parsed_data = normalize_blueprint_payload(parsed_data)
            blueprint = PPTFinalBlueprint.model_validate(parsed_data)
            for idx, canvas in enumerate(blueprint.canvases):
                writer_slide = writer_output_json["slides"][idx]
                if writer_slide.get("layout_type") == "svg":
                    for el in canvas.elements:
                        if el.type == "svg":
                            el.svg_code = writer_slide.get("svg_code") 
                            state.set("blueprint", blueprint.model_dump())
            return blueprint

        except Exception as e:
            print(f"[Layout Agent] 尝试自动修复格式失败: {e}，将尝试原通道校验作为兜底...")
            parsed_data = parse_json_object(blueprint_json)
            blueprint = PPTFinalBlueprint.model_validate(normalize_blueprint_payload(parsed_data))
            state.set("blueprint", blueprint.model_dump())
            return blueprint

    return design_layout