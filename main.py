import os
import json
import re
import traceback
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from agents.planner import get_planner_agent
from agents.writer import get_writer_agent
from agents.review import review_ppt_outline
from agents.image_generator import get_image_generator 
from agents.template import get_layout_agent
from renderer import PPTXRenderer
from config.schema import PPTOutline, PPTFinalBlueprint
from config.debug import save_agent_log 
from config.state import PPTGenerationState


def _render_offline_demo(user_topic: str, filename: str) -> str:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt

    output_path = os.path.abspath(filename)
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    lines = [line.strip() for line in user_topic.splitlines() if line.strip()]
    title = lines[0] if lines else "PPT 生成演示"
    subtitle = "本地离线演示模式" if len(lines) <= 1 else lines[1]
    bullets = []
    for line in lines[2:]:
        if line.startswith("要求") or line.startswith("提示"):
            continue
        bullets.append(line)
    if not bullets:
        bullets = ["主题已接收", "可继续扩展内容", "适配网页生成流程"]

    prs = Presentation()
    blank_layout = prs.slide_layouts[6]

    # Cover slide
    cover = prs.slides.add_slide(blank_layout)
    cover.background.fill.solid()
    cover.background.fill.fore_color.rgb = RGBColor(0xF8, 0xFA, 0xFC)
    title_box = cover.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(1.2))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x11, 0x1F, 0x2E)
    p.alignment = PP_ALIGN.LEFT

    subtitle_box = cover.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(10.5), Inches(1.0))
    tf2 = subtitle_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = subtitle
    p2.font.size = Pt(16)
    p2.font.color.rgb = RGBColor(0x65, 0x6B, 0x7A)

    content = prs.slides.add_slide(blank_layout)
    content.background.fill.solid()
    content.background.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    box = content.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.5), Inches(5.6))
    tf3 = box.text_frame
    tf3.word_wrap = True
    tf3.margin_left = Inches(0.08)
    tf3.margin_right = Inches(0.08)
    title_p = tf3.paragraphs[0]
    title_p.text = "生成内容摘要"
    title_p.font.size = Pt(20)
    title_p.font.bold = True
    title_p.font.color.rgb = RGBColor(0x0F, 0x5B, 0xA4)

    for bullet in bullets[:5]:
        p = tf3.add_paragraph()
        p.text = f"• {bullet}"
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(0x33, 0x3F, 0x4B)
        p.level = 0
        p.space_after = Pt(8)

    summary = prs.slides.add_slide(blank_layout)
    summary.background.fill.solid()
    summary.background.fill.fore_color.rgb = RGBColor(0xF8, 0xFA, 0xFC)
    box2 = summary.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.2), Inches(4.2))
    tf4 = box2.text_frame
    tf4.word_wrap = True
    p4 = tf4.paragraphs[0]
    p4.text = "网页端已成功接入 PPT 生成流程"
    p4.font.size = Pt(22)
    p4.font.bold = True
    p4.font.color.rgb = RGBColor(0x0F, 0x5B, 0xA4)
    p5 = tf4.add_paragraph()
    p5.text = "如配置了 DASHSCOPE_API_KEY，系统会继续走完整的 Planner / Writer / Layout 智能体链路。"
    p5.font.size = Pt(16)
    p5.font.color.rgb = RGBColor(0x65, 0x6B, 0x7A)

    prs.save(output_path)
    return output_path


def run_incremental_modify(
    user_topic: str,
    page_number: int,
    filename: str = "output.pptx",
    state: Optional[PPTGenerationState] = None,
    requirements: str = "",
    run_dir: Optional[str] = None,
):
    """单页增量修改（兼容入口）：转调多页增量管线。"""
    return run_incremental_modify_pages(
        user_topic,
        page_numbers=[page_number],
        filename=filename,
        state=state,
        requirements=requirements,
        run_dir=run_dir,
    )


def run_incremental_modify_pages(
    user_topic: str,
    page_numbers: Optional[List[int]],
    filename: str = "output.pptx",
    state: Optional[PPTGenerationState] = None,
    requirements: str = "",
    run_dir: Optional[str] = None,
):
    """
    页面级增量修改管线（支持多页 / 全局）：仅对指定页集合调用 Writer / 生图 / Layout，
    其余页面完全沿用缓存的大纲与物理蓝图，最后整体重渲染导出。

    :param user_topic: 用户本轮输入（主题 + 最新修改要求）
    :param page_numbers: 目标页码列表（从 1 开始）；为空或 None 表示全局（全部页）
    :param filename: 最终导出的 pptx 文件路径
    :param state: 携带缓存 outline/blueprint 的会话状态
    :param requirements: 用户最新要求
    :param run_dir: 统一的大时间戳文件夹路径；为空时回退 debug_logs/run_{ts}
    """
    api_key = os.environ.get("DASHSCOPE_API_KEY")
    if not api_key:
        print("[Incremental] 未检测到 DASHSCOPE_API_KEY，回退全量（离线演示）管线。")
        return run_ppt_system(user_topic, filename=filename, state=state, requirements=requirements, run_dir=run_dir)

    if state is None:
        state = PPTGenerationState(topic=user_topic, requirements=requirements)
    else:
        state.topic = user_topic
        state.requirements = requirements

    cached_outline = state.get("outline")
    cached_blueprint = state.get("blueprint")
    slides = cached_outline.get("slides") if isinstance(cached_outline, dict) else None
    canvases = cached_blueprint.get("canvases") if isinstance(cached_blueprint, dict) else None

    def _fallback(reason: str) -> Optional[str]:
        print(f"[Incremental] {reason}，回退全量生成管线。")
        state.append_progress(f"Incremental: {reason}，已回退全量生成")
        # 恢复缓存，避免中途失败污染历史状态
        if isinstance(cached_outline, dict) and cached_outline:
            state.set("outline", cached_outline)
        if isinstance(cached_blueprint, dict) and cached_blueprint:
            state.set("blueprint", cached_blueprint)
        return run_ppt_system(user_topic, filename=filename, state=state, requirements=requirements, run_dir=run_dir)

    if not isinstance(slides, list) or not slides or not isinstance(canvases, list) or not canvases:
        return _fallback("缓存的大纲或物理蓝图缺失/非法")
    total = len(slides)
    if total != len(canvases):
        return _fallback(f"大纲与蓝图页数不对齐（{total} vs {len(canvases)}）")

    # 解析目标页集合：None/空 → 全局（全部页）
    targets = list(page_numbers) if page_numbers else list(range(1, total + 1))
    targets = sorted(set(targets))
    if not targets:
        return _fallback("目标页集合为空")
    if any(p < 1 or p > total for p in targets):
        return _fallback(f"页码 {targets} 超出范围（当前共 {total} 页）")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_log_dir = os.path.join(run_dir, "run") if run_dir else f"debug_logs/run_{timestamp}"
    topic = cached_outline.get("topic") or state.topic

    state.append_progress(f"Incremental: 开始 {', '.join(map(str, targets))} 页增量修改")
    print(f"\n[Incremental] 目标页 {targets}（共 {total} 页）增量修改启动，其余页面沿用缓存。")

    try:
        writer = get_writer_agent(api_key=api_key, model_name="qwen3.8-max")
        image_generator = get_image_generator(api_key=api_key, model_name="wan2.7-t2v-2026-06-12")
        layout_designer = get_layout_agent(api_key=api_key, model_name="qwen3.8-max")

        merged_slides = list(slides)
        merged_canvases = list(canvases)

        for page_number in targets:
            page_idx = page_number - 1
            target_slide = slides[page_idx] if isinstance(slides[page_idx], dict) else {}

            # 邻页主题色：注入输入 payload，约束单页修改的色彩家族连续性
            neighbor_themes = [
                slides[i].get("theme")
                for i in (page_idx - 1, page_idx + 1)
                if 0 <= i < total and isinstance(slides[i], dict)
            ]
            neighbor_themes = [t for t in neighbor_themes if isinstance(t, dict)]

            # ---- 文本/主题层：仅目标页改写（design 也走 Writer，但锁文字只改 theme） ----
            single_payload = {"topic": topic, "slides": [target_slide]}
            if neighbor_themes:
                single_payload["neighbor_themes"] = neighbor_themes
            updated_outline = writer(state, single_payload)
            if len(updated_outline.slides) != 1:
                print(f"[Incremental] 警告：Writer 返回 {len(updated_outline.slides)} 页，仅取第 1 页作为目标页。")
            updated_slide = updated_outline.slides[0].model_dump()

            # ---- 图像层：仅目标页生图（SVG 页跳过，与全量管线策略一致） ----
            has_svg = updated_slide.get("layout_type") == "svg" or (updated_slide.get("svg_code") or "").strip()
            if has_svg:
                updated_slide["image_path"] = None
                state.append_progress(f"Incremental: 第 {page_number} 页为 SVG 页，跳过生图")
            elif (updated_slide.get("image_prompt") or "").strip():
                state.append_progress(f"Incremental: 仅生成第 {page_number} 页插图")
                theme = updated_slide.get("theme") or {}
                enhanced_prompt = f"""
                    {updated_slide.get('image_prompt')}。画面主色调采用：{theme.get('primary')}，辅助高亮色采用：{theme.get('secondary')}。
                    这是一张 16:9 演示文稿辅助插图，只用于侧边、角落或独立插图区点缀，不作为背景图或文字衬底。
                    画面需要主体清晰、边缘可裁切、构图简洁，避免中心堆满复杂纹理。
                    严禁生成任何具体的英文字母、汉字或文本，严禁生成复杂统计指标、柱状图等需要精确阅读的图表。
                    避免高对比杂乱纹理、暗部噪点和会干扰正文阅读的复杂背景；整体清爽、克制。
                    """
                local_img_path = image_generator(
                    prompt=enhanced_prompt,
                    size="1280x720",
                    output_dir=run_log_dir,
                    filename=f"slide_accent_{page_idx}.png"
                )
                updated_slide["image_path"] = local_img_path
            else:
                updated_slide["image_path"] = None

            # ---- 物理层：仅目标页重新排版 ----
            state.append_progress(f"Incremental: 生成第 {page_number} 页物理蓝图")
            layout_payload = {"topic": topic, "slides": [updated_slide]}
            if neighbor_themes:
                layout_payload["neighbor_themes"] = neighbor_themes
            page_blueprint = layout_designer(state, layout_payload)
            if not page_blueprint.canvases:
                raise ValueError(f"Layout Agent 未返回第 {page_number} 页有效画布")
            if len(page_blueprint.canvases) != 1:
                print(f"[Incremental] 警告：Layout 返回 {len(page_blueprint.canvases)} 个画布，仅取第 1 个。")
            new_canvas = page_blueprint.canvases[0]

            # SVG 回填（与全量管线 template.py 的对齐逻辑保持一致）
            if updated_slide.get("layout_type") == "svg":
                for el in new_canvas.elements:
                    if el.type == "svg":
                        el.svg_code = updated_slide.get("svg_code")

            # ---- 拼接目标页，其余页保持缓存原样 ----
            merged_slides[page_idx] = updated_slide
            merged_canvases[page_idx] = new_canvas.model_dump()

        merged_blueprint = PPTFinalBlueprint.model_validate({"canvases": merged_canvases})

        state.set("outline", {"topic": topic, "slides": merged_slides})
        state.set("blueprint", merged_blueprint.model_dump())
        save_agent_log({"topic": topic, "slides": merged_slides}, "incremental_outline.json", run_log_dir)
        save_agent_log(merged_blueprint.model_dump(), "incremental_blueprint.json", run_log_dir)

        renderer = PPTXRenderer(merged_blueprint)
        renderer.render_all(filename)
        state.append_progress("Incremental: 完成，PPT 导出成功")
        print(f" [Incremental] 目标页 {targets} 增量修改闭环，成功导出至: {filename}")
        return filename
    except Exception as e:
        print(f"\n[Incremental] 增量修改发生崩溃！故障报告: {e}")
        traceback_text = traceback.format_exc()
        os.makedirs(run_log_dir, exist_ok=True)
        save_agent_log({"error": str(e), "pages": targets}, "error.json", run_log_dir)
        error_txt_path = os.path.join(run_log_dir, "error.txt")
        with open(error_txt_path, "w", encoding="utf-8") as f:
            f.write(traceback_text)
        return _fallback("增量修改过程发生异常")


def run_ppt_system(
    user_topic: str, 
    filename: str = "output.pptx",
    from_content_json: Optional[str] = None,
    from_blueprint_json: Optional[str] = None,
    state: Optional[PPTGenerationState] = None,
    requirements: str = "",
    run_dir: Optional[str] = None,
):
    """
    PPT 自动化高阶管线总线
    
    :param user_topic: PPT 主题描述
    :param filename: 最终导出的 pptx 文件路径
    :param from_content_json: 可选，本地缓存的 PPTOutline JSON 路径。传入后将跳过大纲生产和生图阶段。
    :param from_blueprint_json: 可选，本地缓存的 PPTFinalBlueprint JSON 路径。传入后将直接渲染。
    :param run_dir: 统一的大时间戳文件夹路径；为空时回退 debug_logs/run_{ts}
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_log_dir = os.path.join(run_dir, "run") if run_dir else f"debug_logs/run_{timestamp}"
    

    if from_blueprint_json:
        print(f"\n 检测到物理蓝图缓存注入：{from_blueprint_json}")
        if not os.path.exists(from_blueprint_json):
            print(f"错误：指定的蓝图缓存文件不存在！")
            return None
        
        try:
            with open(from_blueprint_json, "r", encoding="utf-8") as f:
                blueprint_dict = json.load(f)
            
            # 反序列化为 Pydantic 强校验对象
            final_blueprint = PPTFinalBlueprint.model_validate(blueprint_dict)
            print("物理蓝图反序列化成功！")
            
            # 直接移交哑巴渲染器
            renderer = PPTXRenderer(final_blueprint)
            renderer.render_all(filename)
            return filename
        except Exception as e:
            print(f"蓝图注入并渲染失败: {e}")
            return None

    api_key = os.environ.get("DASHSCOPE_API_KEY")
    if not api_key:
        print("未检测到 DASHSCOPE_API_KEY，使用离线演示模式生成基础 PPT。")
        return _render_offline_demo(user_topic, filename)

    if state is None:
        state = PPTGenerationState(topic=user_topic, requirements=requirements)
    else:
        state.topic = user_topic
        state.requirements = requirements
    if not state.messages or state.messages[-1].get("role") != "user":
        state.append_message("user", user_topic if not requirements else f"{user_topic}\n要求：{requirements}")
    final_content = None 

    if from_content_json:
        print(f"\n检测到文本语义大纲缓存注入：{from_content_json}")
        if not os.path.exists(from_content_json):
            print(f"错误：指定的内容缓存文件不存在！")
            return None
        
        try:
            with open(from_content_json, "r", encoding="utf-8") as f:
                content_dict = json.load(f)
            final_content = PPTOutline.model_validate(content_dict)
            print("语义大纲反序列化成功！")
        except Exception as e:
            print(f"内容大纲注入失败: {e}")
            return None
            
    else:
        try:
            planner = get_planner_agent(api_key=api_key, model_name="qwen3.8-max")
            writer = get_writer_agent(api_key=api_key, model_name="qwen3.8-max")
            image_generator = get_image_generator(api_key=api_key, model_name="qwen-image-3.0")                
            current_feedback = None
            max_retries = 3
            attempt = 0
            
            # 内容生产与质检闭环
            while attempt < max_retries:
                attempt += 1
                progress_msg = f"第 {attempt} 轮智能体协同" 
                state.append_progress(progress_msg)
                print(f"\n {progress_msg} ")
                
                raw_outline = planner(state, feedback_list=current_feedback)
                save_agent_log(raw_outline.model_dump(), f"raw_outline_attempt_{attempt}.json", run_log_dir)
                compiled_content = writer(state, raw_outline)
                save_agent_log(compiled_content.model_dump(), f"compiled_content_attempt_{attempt}.json", run_log_dir)
                
                # 检查内容合规性
                is_passed, errors = review_ppt_outline(compiled_content)
                
                if is_passed:
                    print(f"[Reviewer] 内容终审通过！准备唤醒生图和物理画布排版。")
                    final_content = compiled_content
                    save_agent_log(final_content.model_dump(), f"final_content_passed.json", run_log_dir)
                    break
                else:
                    print(f"[Reviewer] 质检失败！已派发回上一级。")
                    current_feedback = errors  
                    
            if not final_content:
                print(f"\n系统提示：未能在限制轮次内完美对齐，采用最后一轮输出保底。")
                final_content = compiled_content

            # 调用图像模型生成图并写入路径
            print("\n正在遍历各幻灯片，触发图片生成管道...")
            for i, slide in enumerate(final_content.slides):
                if slide.layout_type == "svg" or (slide.svg_code and slide.svg_code.strip()):
                    slide.image_path = None
                    state.append_progress(f"Images: 第 {i + 1} 页包含 SVG，跳过本地生图")
                    print(f"[Pipeline] 检测到第 {i} 页含有原生高级SVG，已跳过本地物理生图，规避视觉冲突。")
                    continue
                    
                if slide.image_prompt and slide.image_prompt.strip():
                    state.append_progress(f"Images: 生成第 {i + 1} 页插图")
                    enhanced_prompt = f"""
                    {slide.image_prompt}。画面主色调采用：{slide.theme.primary}，辅助高亮色采用：{slide.theme.secondary}。
                    这是一张 16:9 演示文稿辅助插图，只用于侧边、角落或独立插图区点缀，不作为背景图或文字衬底。
                    画面需要主体清晰、边缘可裁切、构图简洁，避免中心堆满复杂纹理。
                    严禁生成任何具体的英文字母、汉字或文本，严禁生成复杂统计指标、柱状图等需要精确阅读的图表。
                    避免高对比杂乱纹理、暗部噪点和会干扰正文阅读的复杂背景；整体清爽、克制。
                    """
                    local_img_path = image_generator(
                        prompt=enhanced_prompt,
                        size="1280x720",
                        output_dir=run_log_dir,
                        filename=f"slide_accent_{i}.png"
                    )
                    slide.image_path = local_img_path
            
            save_agent_log(final_content.model_dump(), f"final_content_with_images.json", run_log_dir)
            state.append_progress("Images: 生成并写入插图路径")
        except Exception as e:
            print(f"\n内容生成或图像生成阶段发生崩溃！故障报告: {e}")
            traceback_text = traceback.format_exc()
            os.makedirs(run_log_dir, exist_ok=True)
            save_agent_log({"error": str(e)}, "error.json", run_log_dir)
            error_txt_path = os.path.join(run_log_dir, "error.txt")
            with open(error_txt_path, "w", encoding="utf-8") as f:
                f.write(traceback_text)
            state.append_progress("错误：已转离线演示")
            print("已切换为离线演示模式，继续生成基础 PPT。")
            return _render_offline_demo(user_topic, filename)

    # 视觉排版设计与物理画布渲染管线
    # 逐页排版：每次仅向 Layout Agent 提交单页语义大纲，避免一次性生成全部页物理蓝图，
    # 导致单次请求输出体积过大 / 生成耗时过长，触发 MaaS 网关主动断连（Connection error）。
    try:
        layout_designer = get_layout_agent(api_key=api_key, model_name="qwen3.8-max")
        writer_dump_data = final_content.model_dump()
        slides = writer_dump_data.get("slides") or []
        total = len(slides)
        if not total:
            raise ValueError("语义大纲不含任何页面，无法排版")
        topic = writer_dump_data.get("topic") or state.topic

        state.append_progress(f"Rendering: 开始逐页布局渲染（共 {total} 页）")
        merged_canvases = []
        for i, slide in enumerate(slides):
            state.append_progress(f"Rendering: 排版第 {i + 1}/{total} 页")
            print(f" [Layout] 逐页排版第 {i + 1}/{total} 页")
            single_payload = {"topic": topic, "slides": [slide]}
            # 邻页主题色：约束单页配色的色彩家族连续性（与增量管线一致）
            neighbor_themes = [
                slides[j].get("theme")
                for j in (i - 1, i + 1)
                if 0 <= j < total and isinstance(slides[j], dict)
            ]
            neighbor_themes = [t for t in neighbor_themes if isinstance(t, dict)]
            if neighbor_themes:
                single_payload["neighbor_themes"] = neighbor_themes

            page_blueprint = layout_designer(state, single_payload)
            if not page_blueprint.canvases:
                raise ValueError(f"Layout Agent 未返回第 {i + 1} 页有效画布")
            if len(page_blueprint.canvases) != 1:
                print(f" [Layout] 警告：第 {i + 1} 页返回 {len(page_blueprint.canvases)} 个画布，仅取第 1 个。")
            merged_canvases.append(page_blueprint.canvases[0].model_dump())

        final_blueprint = PPTFinalBlueprint.model_validate({"canvases": merged_canvases})
        # 显式写回全量蓝图，确保落盘快照（_persist_generation_assets 依赖 state["blueprint"]）
        state.set("blueprint", final_blueprint.model_dump())
        save_agent_log(final_blueprint.model_dump(), f"final_blueprint.json", run_log_dir)
        renderer = PPTXRenderer(final_blueprint)
        renderer.render_all(filename)
        state.append_progress("Rendering: 完成，PPT 导出成功")
        print(f" 幻灯片系统闭环，成功导出至: {filename}")
    except Exception as e:
        print(f" 视觉排版或物理渲染阶段发生崩溃！故障报告: {e}")
        traceback_text = traceback.format_exc()
        os.makedirs(run_log_dir, exist_ok=True)
        save_agent_log({"error": str(e)}, "error.json", run_log_dir)
        error_txt_path = os.path.join(run_log_dir, "error.txt")
        with open(error_txt_path, "w", encoding="utf-8") as f:
            f.write(traceback_text)
        state.append_progress("错误：已转离线演示")
        print("已切换为离线演示模式，继续生成基础 PPT。")
        return _render_offline_demo(user_topic, filename)

    return filename


if __name__ == "__main__":
    # 场景一：纯粹调试本地 python-pptx 形状绘制、字体颜色、图层覆盖
    #run_ppt_system(
    #     user_topic="通俗易懂地解释词向量（Word Embedding）的基本概念", 
    #     filename="output.pptx",
    #     from_blueprint_json=""
    #)
    #场景二：跳过长文本编写与图片生成，微调 Layout Agent 的网格算框 Prompt 策略
    #run_ppt_system(
        #user_topic="大语言模型的基本原理", 
        #filename="Coffee.pptx",
        #from_content_json=""
    #)
    # 场景三：常规全量端到端生成
    #run_ppt_system("大语言模型的基本原理", "LLM.pptx")
    run_ppt_system("通俗易懂地解释词向量（Word Embedding）的基本概念", "WE.pptx")
    #run_ppt_system("绘制 SVG 流程图，展示从一颗咖啡豆到一杯咖啡的完整生产链（种植、采摘、烘焙、研磨、冲煮）", "COFFEE.pptx")
    #run_ppt_system("YouTube has 10 times more videos than TikTok, TikTok has 2 times more than Kuaishou", "YK.pptx")
