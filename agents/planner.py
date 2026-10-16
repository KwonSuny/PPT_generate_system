import json
from typing import List, Optional
from config.content import identified_outline
from config.schema import PagePolicy
from config.sanitizers import normalize_outline_payload, parse_completion_payload
from config.state import PPTGenerationState
from config.llm_client import create_client, chat_completion_with_retry, resolve_model_name, model_request_options


def get_planner_agent(api_key: str, model_name: Optional[str] = None):
    model_name = model_name or resolve_model_name("planner")
    client = create_client(api_key)

    def plan_ppt(state: PPTGenerationState, feedback_list: Optional[List[str]] = None):
        policy = PagePolicy.model_validate(state.get("page_policy") or {})
        system_prompt = """
<role>
你是兼具叙事编辑、信息设计和艺术指导能力的演示文稿规划师。你的交付是让观众看懂并记住主题的整套设计，而不是把文章切成相似的文字页。
</role>
<task>
先识别受众、讲述目的、已有事实、页数与风格，规划每页唯一的中心信息。使用开场、解释、对照、过程、证据与收束形成叙事节奏；不为凑页数重复标题和段落。
风格优先级：用户明确说明 > 根据主题和用途识别 > 简洁汇报。不得强制深色科技风、渐变或卡片。
根结构：schema_version=2、topic、slides。页数指总页数，包含封面；一页请求不必添加封面。
每页保留：id、title、punchline、bullet_points、bullet_ids、full_content、layout_type、theme、table、diagram、svg_code、icons、image_prompt、image_path。
id 使用 slide-1 等稳定标识；bullet_ids 与 bullet_points 一一对应，使用 point-1 等标识。
theme 必须有 primary、secondary、content、background 四个 #RRGGBB 色值。
</task>
<art_direction>
用户明确风格优先于旧版、邻页与默认配色。为本主题选择主色、辅助色和中性色，保持色族一致，但一致不等于每页白底蓝字。
允许同一套内使用主色深底开场/章节、浅色解释页、强调色结论页；背景可使用低饱和色面、侧栏、页眉或几何构图。根据语义安排，不要机械轮换颜色或给所有内容加同款卡片。
文字可通过字号、粗细、强调色和留白形成主次：标题表达结论，正文解释证据。深色背景配浅色正文，浅色背景配深色正文；primary 必须能在该页背景上清晰阅读。
构图至少考虑图文左右、上下主视觉、对比列、过程带、时间轴和大结论。三页以上不连续使用同一种构图。给关键图示足够面积，不能每页都出现三条同样长的段落。
可参考科技主题的藏蓝/青绿、自然主题的松绿/赭色、学术主题的墨色/朱红等关联，但这些是灵感，不是必须套用的模板。不要把“汇报”理解为只能白底蓝字。
</art_direction>
<content_and_capacity>
layout_type 选择：title_page（简洁封面）、standard_text（说明）、column（对比）、timeline（历史）、table（行列数据）、diagram（简单流程或架构）、svg（复杂构图）。
用户明确要求表格时，至少规划一页 table；表格可用于步骤、特征等定性对比，不要求虚构数字。默认简单流程和关系选 diagram，数据表始终用原生 table。
用户明确要求内联 SVG 表达时，规划适合主题的矢量插图/总览并交由 Writer 生成 svg_code，可附在封面、说明或表格页面；不能把这一要求忽略为仅仅禁止照片。用户明确指定 SVG 流程时遵循指定，其他简单流程保留原生 diagram。
table、diagram 必须留 null，由 Writer 根据原始资料补全；不要自定义 headers/rows 等结构。其他扩展字段使用 [] 或 null。
每页通常1–3条要点，必要时4条；要点是10–24字的短标题，不应提前写成长正文，每条硬上限50字。full_content 留 []。
图表为主的页面只保留必要解释；避免在复杂图示外再堆三四条长标题和长正文。结构化页面可不列要点。用户提供的事实、数值与来源必须完整安置，不能为适配布局删掉关键证据。
标题直接说明主题或有证据的结论；不得用虚构数据、泛泛口号、页面设计说明填充内容。
保留用户数据的数值、精度、单位、符号、时间、比较条件和来源。缺少数据时说明待补充，不编造对比数值。
</content_and_capacity>
<output>
只输出合法 JSON，不输出 Markdown 或 schema 定义。提交前检查：页数、稳定ID、各页中心信息、背景与文字对比、构图节奏、内容密度。
</output>
"""
        if feedback_list:
            state.feedback = feedback_list
        prompt = (
            f"主题：{state.topic}\n要求：{state.requirements}\n"
            f"总页数必须在 {policy.min_slides}–{policy.max_slides} 页之间，包含封面。"
        )
        if state.should_use_incremental_mode():
            prompt += "基于现有内容优化，遵守本次修改范围。"
        if feedback_list:
            prompt += "\n需要修正的问题：" + json.dumps(feedback_list, ensure_ascii=False)
        state.append_progress("Planner: 生成大纲中")
        response = chat_completion_with_retry(
            client, model=model_name, messages=state.build_messages(system_prompt, prompt, role="planner"),
            response_format={"type": "json_object"},
            **model_request_options('planner', model_name, search=True),
        )
        payload = normalize_outline_payload(parse_completion_payload(response, state, "Planner"))
        # Planner 只拥有页面骨架，扩展数据由 Writer 按统一契约生成。
        for slide in payload["slides"]:
            slide["table"] = slide["diagram"] = None
        outline = identified_outline(payload)
        state.set("outline", outline.model_dump())
        return outline

    return plan_ppt
