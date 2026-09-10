"""
增量修改管线集成测试（不依赖真实 LLM API）。

覆盖点：
1. run_incremental_modify 正确拼接：页数不变、目标页替换、其余页保持缓存、仅目标页生图
2. 页码越界时回退全量管线（离线演示模式）
3. Flask 路由：增量条件满足时走 incremental_page 分支并落盘资产快照

运行：PATH 前置 F:/conda/Library/bin 后用 conda python 执行。
"""
import base64
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
os.environ["DASHSCOPE_API_KEY"] = "fake-key-for-local-test"

from pptx import Presentation

from config.schema import (
    CanvasElement,
    PPTFinalBlueprint,
    PPTOutline,
    SlideCanvas,
    SlideElement,
    TextParagraph,
    ColorPalette,
)
from config.state import PPTGenerationState
import main as main_module

TEST_OUT_DIR = Path(__file__).resolve().parent / "debug_logs" / "incremental_test"
TEST_OUT_DIR.mkdir(parents=True, exist_ok=True)

# 1x1 透明 PNG，用于假生图 Agent
TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)

THEME = ColorPalette(primary="#1E3A8A", secondary="#10B981", content="#64748B", background="#F8FAFC")


def make_slide(title, punchline, bullets, image_prompt=None):
    return SlideElement(
        title=title,
        punchline=punchline,
        bullet_points=bullets,
        full_content=bullets,
        theme=THEME,
        layout_type="standard_text",
        image_prompt=image_prompt,
    )


def make_canvas(title):
    return SlideCanvas(
        background_color="#FFFFFF",
        elements=[
            CanvasElement(
                type="textbox",
                left=0.8,
                top=0.8,
                width=11.5,
                height=1.2,
                paragraphs=[TextParagraph(text=title, font_size=24, bold=True)],
            ),
            CanvasElement(
                type="textbox",
                left=0.8,
                top=2.2,
                width=11.5,
                height=4.0,
                paragraphs=[TextParagraph(text=f"{title}-正文缓存内容", font_size=14)],
            ),
        ],
    )


def build_cached_state():
    """构造带 3 页缓存大纲 + 3 页缓存蓝图的会话状态。"""
    slides = [
        make_slide("第1页原始标题", "页面一摘要", ["要点A1", "要点A2"]),
        make_slide("第2页原始标题", "页面二摘要", ["要点B1", "要点B2"], image_prompt="原始配图意境"),
        make_slide("第3页原始标题", "页面三摘要", ["要点C1", "要点C2"]),
    ]
    outline = PPTOutline(topic="测试主题", slides=slides)
    blueprint = PPTFinalBlueprint(canvases=[make_canvas(s.title) for s in slides])
    state = PPTGenerationState(topic="测试主题", requirements="把第2页标题改掉")
    state.set("outline", outline.model_dump())
    state.set("blueprint", blueprint.model_dump())
    return state, outline, blueprint


CALLS = {"writer": 0, "layout": 0, "image": 0, "planner": 0}


def install_fake_agents():
    def fake_get_writer_agent(api_key, model_name="x"):
        def writer(state, payload):
            CALLS["writer"] += 1
            assert isinstance(payload, dict) and len(payload["slides"]) == 1, "Writer 应只收到单页输入"
            src = payload["slides"][0]
            if state.modify_scope == "design":
                # 模拟 lock design：只改 theme，保持文本（title/punchline/bullet/full_content 不变）
                new_theme = ColorPalette(primary="#000000", secondary="#111111", content="#222222", background="#FFFFFF")
                updated = SlideElement(
                    title=src.get("title"),
                    punchline=src.get("punchline"),
                    bullet_points=src.get("bullet_points") or [],
                    full_content=src.get("full_content") or [],
                    theme=new_theme,
                    layout_type=src.get("layout_type") or "standard_text",
                    image_prompt=src.get("image_prompt"),
                )
            else:
                updated = SlideElement(
                    title="第2页-增量更新标题",
                    punchline="增量修改后的摘要",
                    bullet_points=["增量要点1", "增量要点2"],
                    full_content=["增量正文1", "增量正文2"],
                    theme=THEME,
                    layout_type="standard_text",
                    image_prompt="新的配图意境",
                )
            # 模拟真实 Writer 的副作用：把 state.outline 覆盖为单页版本
            state.set("outline", {"topic": payload.get("topic", "测试主题"), "slides": [updated.model_dump()]})
            return PPTOutline(topic=payload.get("topic", "测试主题"), slides=[updated])

        return writer

    def fake_get_layout_agent(api_key, model_name="x"):
        def layout(state, writer_output_json):
            CALLS["layout"] += 1
            assert len(writer_output_json["slides"]) == 1, "Layout 应只收到单页输入"
            slide = writer_output_json["slides"][0]
            elements = [
                CanvasElement(
                    type="textbox",
                    left=0.8,
                    top=0.8,
                    width=11.5,
                    height=1.2,
                    paragraphs=[TextParagraph(text=slide["title"], font_size=24, bold=True)],
                ),
                CanvasElement(
                    type="textbox",
                    left=0.8,
                    top=2.2,
                    width=8.0,
                    height=3.0,
                    paragraphs=[TextParagraph(text=" | ".join(slide.get("full_content") or []), font_size=14)],
                ),
            ]
            if slide.get("image_path"):
                elements.append(
                    CanvasElement(
                        type="image",
                        left=9.5,
                        top=2.2,
                        width=3.0,
                        height=2.0,
                        image_path=slide["image_path"],
                    )
                )
            # 模拟真实 Layout 的副作用
            state.set("blueprint", {"canvases": [SlideCanvas(background_color="#FFFFFF", elements=elements).model_dump()]})
            return PPTFinalBlueprint(canvases=[SlideCanvas(background_color="#FFFFFF", elements=elements)])

        return layout

    def fake_get_image_generator(api_key, model_name="x"):
        def image_gen(prompt, *args, size="1280x720", output_dir="output_images", filename=None, **kwargs):
            CALLS["image"] += 1
            os.makedirs(output_dir, exist_ok=True)
            path = os.path.abspath(os.path.join(output_dir, filename or "fake.png"))
            with open(path, "wb") as f:
                f.write(TINY_PNG)
            return path

        return image_gen

    def fake_get_planner_agent(api_key, model_name="x"):
        def planner(state, feedback_list=None):
            CALLS["planner"] += 1
            raise ValueError("测试环境禁止调用 Planner")

        return planner

    main_module.get_writer_agent = fake_get_writer_agent
    main_module.get_layout_agent = fake_get_layout_agent
    main_module.get_image_generator = fake_get_image_generator
    main_module.get_planner_agent = fake_get_planner_agent


def slide_texts(prs, idx):
    texts = []
    for shape in prs.slides[idx].shapes:
        if hasattr(shape, "text") and shape.text:
            texts.append(shape.text.strip())
    return texts


def test_incremental_splice():
    print("=" * 60)
    print("TEST 1: 增量拼接正确性（第 2 页修改，共 3 页）")
    print("=" * 60)
    install_fake_agents()
    state, outline, blueprint = build_cached_state()
    output = str(TEST_OUT_DIR / "incremental_splice.pptx")

    result = main_module.run_incremental_modify(
        "测试主题\n要求：把第2页标题改掉",
        page_number=2,
        filename=output,
        state=state,
        requirements="把第2页标题改掉",
    )
    assert result == output, f"应返回输出路径，实际: {result}"

    prs = Presentation(output)
    assert len(prs.slides) == 3, f"页数应保持 3，实际 {len(prs.slides)}"

    t1, t2, t3 = (slide_texts(prs, i) for i in range(3))
    assert any("第1页原始标题" in t for t in t1), "第 1 页应保持缓存原样"
    assert any("第3页原始标题" in t for t in t3), "第 3 页应保持缓存原样"
    assert any("第2页-增量更新标题" in t for t in t2), "第 2 页应替换为增量更新内容"
    assert any("增量正文" in t for t in t2), "第 2 页正文应为增量内容"

    merged_outline = state.get("outline")
    assert len(merged_outline["slides"]) == 3, "state.outline 应恢复为 3 页（拼接后）"
    assert merged_outline["slides"][1]["title"] == "第2页-增量更新标题"
    assert merged_outline["slides"][0]["title"] == "第1页原始标题"

    merged_blueprint = state.get("blueprint")
    assert len(merged_blueprint["canvases"]) == 3, "state.blueprint 应为 3 页画布"

    assert CALLS["writer"] == 1, f"Writer 应仅调用 1 次，实际 {CALLS['writer']}"
    assert CALLS["layout"] == 1, f"Layout 应仅调用 1 次，实际 {CALLS['layout']}"
    assert CALLS["image"] == 1, f"生图应仅调用 1 次（仅目标页），实际 {CALLS['image']}"
    assert CALLS["planner"] == 0, "增量管线不应调用 Planner"
    print("PASS: 页数不变 / 目标页替换 / 其余页保持缓存 / 仅目标页生图 / state 正确回写\n")


def test_out_of_range_fallback():
    print("=" * 60)
    print("TEST 2: 页码越界回退全量管线")
    print("=" * 60)
    install_fake_agents()
    state, outline, blueprint = build_cached_state()
    output = str(TEST_OUT_DIR / "fallback_out_of_range.pptx")

    result = main_module.run_incremental_modify(
        "测试主题\n要求：修改第 99 页",
        page_number=99,
        filename=output,
        state=state,
        requirements="修改第 99 页",
    )
    assert result == output, f"回退路径应仍返回输出文件，实际: {result}"
    prs = Presentation(output)
    assert len(prs.slides) == 3, "离线演示模式应生成 3 页"
    restored = state.get("outline")
    assert len(restored["slides"]) == 3, "回退后缓存 outline 应被恢复为 3 页"
    print("PASS: 越界页码触发回退，缓存状态被正确恢复\n")


def test_design_scope_keeps_text():
    print("=" * 60)
    print("TEST 3: modify_design 范围下锁文字、只改 theme")
    print("=" * 60)
    install_fake_agents()
    CALLS["writer"] = 0
    CALLS["image"] = 0
    state, outline, blueprint = build_cached_state()
    state.set_operation("incremental_page", scope="design", page=2)
    output = str(TEST_OUT_DIR / "design_scope.pptx")

    result = main_module.run_incremental_modify(
        "测试主题\n要求：第2页换个配色",
        page_number=2,
        filename=output,
        state=state,
        requirements="第2页换个配色",
    )
    assert result == output
    prs = Presentation(output)
    assert len(prs.slides) == 3
    t2 = slide_texts(prs, 1)
    assert any("第2页原始标题" in t for t in t2), "design 模式下第 2 页文本应保持缓存原样"
    assert CALLS["writer"] == 1, f"design 范围应调用 Writer（lock design 只改 theme），实际 {CALLS['writer']}"
    assert CALLS["image"] == 1, "design 范围仍应为目标页重生成配图（image_prompt 存在）"
    merged_outline = state.get("outline")
    assert merged_outline["slides"][1]["theme"]["primary"] == "#000000", "design 模式 theme 应被更新"
    print("PASS: design 范围锁文字，文本保持缓存、theme 更新\n")


def test_incremental_multi_page():
    print("=" * 60)
    print("TEST 4: 多页增量修改（第 2、3 页）")
    print("=" * 60)
    install_fake_agents()
    CALLS["writer"] = 0
    CALLS["layout"] = 0
    CALLS["image"] = 0
    state, outline, blueprint = build_cached_state()
    output = str(TEST_OUT_DIR / "multi_page.pptx")

    result = main_module.run_incremental_modify_pages(
        "测试主题\n要求：改第2、3页",
        page_numbers=[2, 3],
        filename=output,
        state=state,
        requirements="改第2、3页",
    )
    assert result == output, f"应返回输出路径，实际: {result}"
    prs = Presentation(output)
    assert len(prs.slides) == 3, "页数应保持 3"

    t1, t2, t3 = (slide_texts(prs, i) for i in range(3))
    assert any("第1页原始标题" in t for t in t1), "第 1 页应保持缓存原样"
    assert any("增量更新标题" in t for t in t2), "第 2 页应被修改"
    assert any("增量更新标题" in t for t in t3), "第 3 页应被修改"
    assert CALLS["writer"] == 2, f"Writer 应调用 2 次（第2、3页），实际 {CALLS['writer']}"
    assert CALLS["layout"] == 2, f"Layout 应调用 2 次，实际 {CALLS['layout']}"
    assert CALLS["image"] == 2, f"生图应调用 2 次（两页均有 image_prompt），实际 {CALLS['image']}"
    print("PASS: 多页增量修改，目标页替换、其余页冻结\n")


def test_incremental_global():
    print("=" * 60)
    print("TEST 5: 全局增量修改（全部页，跳过 Planner）")
    print("=" * 60)
    install_fake_agents()
    CALLS["writer"] = 0
    CALLS["layout"] = 0
    CALLS["planner"] = 0
    state, outline, blueprint = build_cached_state()
    state.set_operation("incremental_page", scope="text", pages=[1, 2, 3])
    output = str(TEST_OUT_DIR / "global_modify.pptx")

    result = main_module.run_incremental_modify_pages(
        "测试主题\n要求：全部页改文本",
        page_numbers=None,  # 全局
        filename=output,
        state=state,
        requirements="全部页改文本",
    )
    assert result == output
    prs = Presentation(output)
    assert len(prs.slides) == 3
    assert CALLS["writer"] == 3, f"全局应调用 Writer 3 次（每页一次），实际 {CALLS['writer']}"
    assert CALLS["layout"] == 3, f"全局应调用 Layout 3 次，实际 {CALLS['layout']}"
    assert CALLS["planner"] == 0, "全局增量不应调用 Planner"
    print("PASS: 全局增量修改（全部3页，跳过 Planner）\n")


def test_lock_scope_sanitizers():
    print("=" * 60)
    print("TEST 6: sanitizers 字段级锁定（text/design）")
    print("=" * 60)
    from config.sanitizers import normalize_outline_payload
    skeleton = {
        "topic": "t",
        "slides": [{
            "title": "原标题", "punchline": "原摘要", "bullet_points": ["原要点"],
            "full_content": ["原正文"], "layout_type": "column",
            "theme": {"primary": "#111111", "secondary": "#222222", "content": "#333333", "background": "#444444"},
        }],
    }
    # text 锁：文字可改，layout_type/theme 锁定
    raw_text = {"slides": [{"title": "新标题", "bullet_points": ["新要点"], "full_content": ["新正文"],
                            "layout_type": "svg", "theme": {"primary": "#999999"}}]}
    s = normalize_outline_payload(raw_text, skeleton=skeleton, lock_scope="text")["slides"][0]
    assert s["title"] == "新标题", "text 锁应允许改 title"
    assert s["bullet_points"] == ["新要点"], "text 锁应允许改 bullet_points"
    assert s["layout_type"] == "column", "text 锁应保持 layout_type=column"
    assert s["theme"]["primary"] == "#111111", "text 锁应保持 theme.primary"

    # design 锁：文字锁定，theme/layout_type 可改
    raw_design = {"slides": [{"title": "改标题", "bullet_points": ["改要点"], "full_content": ["改正文"],
                              "layout_type": "svg",
                              "theme": {"primary": "#AAAAAA", "secondary": "#BBBBBB", "content": "#CCCCCC", "background": "#DDDDDD"}}]}
    s2 = normalize_outline_payload(raw_design, skeleton=skeleton, lock_scope="design")["slides"][0]
    assert s2["title"] == "原标题", "design 锁应保持 title"
    assert s2["bullet_points"] == ["原要点"], "design 锁应保持 bullet_points"
    assert s2["full_content"] == ["原正文"], "design 锁应保持 full_content"
    assert s2["layout_type"] == "svg", "design 锁应允许 layout_type 改变"
    assert s2["theme"]["primary"] == "#AAAAAA", "design 锁应允许 theme 改变"
    print("PASS: lock_scope text/design 字段级锁定正确\n")


if __name__ == "__main__":
    test_incremental_splice()
    test_out_of_range_fallback()
    test_design_scope_keeps_text()
    test_incremental_multi_page()
    test_incremental_global()
    test_lock_scope_sanitizers()
    print("ALL INCREMENTAL PIPELINE TESTS PASSED")
