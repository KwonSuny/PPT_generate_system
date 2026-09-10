# agents/planner.py
import json
from config.schema import PPTOutline
from typing import List, Optional
from config.sanitizers import normalize_outline_payload, parse_json_object
from config.state import PPTGenerationState
from config.llm_client import create_client, chat_completion_with_retry


def get_planner_agent(api_key: str, model_name: str = "qwen3.7-max"):
    client = create_client(api_key)
    
    def plan_ppt(state: PPTGenerationState, feedback_list: Optional[List[str]] = None) -> PPTOutline:
        system_prompt = """
        # ROLE
        你是一个专业的幻灯片演示专家、首席内容架构师与精确数据结构化专家。你精通顶级咨询公司的 PPT 排版逻辑与金字塔原理。

        # TASK
        根据用户给定的主题，撰写一份 PPT 的大纲，并为每一页幻灯片配置专属的色彩演进方案，拥有多样的形式。输出必须是严格符合 Pydantic Schema 的标准 JSON 大纲。

        # COLOR FAMILY STRATEGY (每页独立的色族演进策略)
        为了彻底打破千篇一律的固定模板，你必须为每一个 `SlideElement`（即每一页幻灯片内部）规划独立的 `theme` (ColorPalette)。
        1. 必须形成“相似色彩家族的系列演进”。例如：若主题是科技类，全套幻灯片的色族基调可以定为“深邃数字流”。
        2. 声明在每一页内部的 `theme` 必须包含以下 4 个 16 进制颜色键名：
           - `primary`: 主色调，用于当前页大标题、核心高亮（如 #005088）
           - `secondary`: 辅助色，用于当前页副标题、图表对比色（如 #11CAA0）
           - `content`: 正文色，用于当前页段落文本（如 #64748B）
           - `background`: 页面背景色（如 #F8FAFC）

        # STRICT OUTPUT RULE
        1. 你必须且只能根据提供的 Pydantic Schema 输出标准的 JSON 数据。
        2. 严禁包含任何 Markdown 标记（如 ```json ）。
        3. 必须完整填充根节点的所有核心字段：`topic`、`total_slides` 和 `slides`。
        4. 对于 `slides` 列表中的每一页，你必须完整填充 `title`, `punchline`, `bullet_points`, `theme`, `layout_type`。
        5. 请将其余字段如 `full_content`, `icons`, `svg_code`, `image_prompt`, `image_path` 统一保持为空列表 `[]` 或 `null`。

        # 正确的 JSON 结构骨架示例（包含根节点与单页内部的 theme）：
        {
            "topic": "演示文稿的主题名称",
            "total_slides": 3,
            "slides": [
                {
                    "title": "单页小标题",
                    "punchline": "核心结论或破局金句",
                    "bullet_points": ["要点骨架1", "要点骨架2"],
                    "full_content": [],
                    "theme": {
                        "primary": "#1E3A8A",
                        "secondary": "#10B981",
                        "content": "#64748B",
                        "background": "#F8FAFC"
                    },
                    "layout_type": "standard_text",
                    "svg_code": null,
                    "icons": [],
                    "image_prompt": null,
                    "image_path": null
                }
            ]
        }
        """
        
        if feedback_list:
            state.feedback = feedback_list

        mode_hint = "请基于已有大纲做增量修改，而不是重新生成全新版本。" if state.should_use_incremental_mode() else "请生成一个新的完整大纲。"

        base_user_constraints = f"""
            # 用户主题
            《{state.topic}》

            # 当前要求
            《{state.requirements}》

            # 全局约束
            - 【总页数死线】：整套 PPT 的总页数（根节点的 total_slides 兼 slides 列表长度）必须严格控制在 1 到 5 页之间！两者数值必须绝对相等！
            - 【增量修改策略】：{mode_hint}
            - 【配色美学】：根据主题调性匹配高级感配色，让每一页的 theme 呈现出高端的视觉演进。且要求每页的 theme 之间必须形成“相似色彩家族的系列演进”，禁止出现突兀的色彩跳跃。
            - 允许每页 1-5 条要点，单条字数精炼。
            
            # SLIDE LAYOUT RULES
            请根据每一页的内容属性，极其严苛地从以下 5 种布局中选择一种（layout_type）：
            1. `title_page`：仅用于第 1 页（封面页）。
            2. `standard_text`：用于核心概念阐述、定义或结论性总结。
            3. `column`：用于事物、方案、优缺点的横向“对比”或“对立”场景。
            4. `timeline`：用于历史演进、技术路线发展、或有明确先后顺序的“步骤/流程”。
            5. `svg`：用于需要绘制专业流程图、架构图，或者展示复杂多维数据表格、对比矩阵的页面。凡是遇到需要以表格（Table）形式呈现非单一趋势的数据场景，务必强行选择 svg 布局，以便通过高级矢量绘图完美闭环，杜绝死板文本排版。
            
            # 内容和数据规范
            - 【字数精炼】：每页的 `bullet_points` 数量控制在 5 条以内。每条长 50 字以内。禁止使用大段落，必须提炼为可直接宣讲的金句、短语。
            - 【SVG 流程图与表格规范】：如果 layout_type 为 `svg`，要求绘制符合规范的 SVG 流程图或现代化精美表格数据矩阵。
            """
        modify_hint = ""
        if state.operation_mode != "new":
            if state.operation_mode == "continue":
                modify_hint = "请基于现有 PPT 进行增量优化，而不是重新生成全新版本。"
            elif state.operation_mode == "modify_text":
                modify_hint = "请仅修改文本内容，不调整原有页面布局风格。"
            elif state.operation_mode == "modify_design":
                modify_hint = "请仅调整设计风格和排版细节，不修改主要文案内容。"
            if state.modify_page is not None:
                modify_hint += f" 重点聚焦 {state.page_label() or ('第 ' + str(state.modify_page) + ' 页')}。"

        if feedback_list:
            feedback_bullets = "\n".join([f"❌ 严重违规项：{err}" for err in feedback_list])
            re_drive_prompt = f"""
            {base_user_constraints}
            {modify_hint}
            【质检驳回警报】
            你上一次生成的 JSON 大纲未能通过 Reviewer Agent 的合规审查。错误如下：
            {feedback_bullets}
            请立即进行深度反思，重新生成一套完全合规的完美的 JSON 数据！
            """
            prompt_text = re_drive_prompt
        else:
            prompt_text = f"{base_user_constraints}\n\n{modify_hint}\n请现在开始规划上述主题，确保完美符合所有规范并以 JSON 格式输出。"

        messages = state.build_messages(system_prompt, prompt_text)

        state.append_progress("Planner: 生成大纲中")
        response = chat_completion_with_retry(
            client,
            model=model_name,
            messages=messages,
            response_format={"type": "json_object"},
            temperature=0.3,
            extra_body={"enable_search": True},
        )
        
        raw_json = response.choices[0].message.content.strip()
        state.append_message("assistant", raw_json)
        try:
            data = parse_json_object(raw_json)
            outline = PPTOutline.model_validate(normalize_outline_payload(data))
            state.set("outline", outline.model_dump())
            return outline
        except Exception as e:
            print(f"[Planner Agent] 输出结构修复失败: {e}，将尝试原始 JSON 通道校验...")
            outline = PPTOutline.model_validate_json(raw_json)
            state.set("outline", outline.model_dump())
            return outline
        
    return plan_ppt