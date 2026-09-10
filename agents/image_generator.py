# agents/image_generator.py
import os
import time
import requests
import dashscope
from dashscope import MultiModalConversation
from typing import Optional
from config.llm_client import create_client, chat_completion_with_retry

def get_image_generator(api_key: str, model_name: str = "qwen-image-2.0-pro-2026-04-22"):
    """
    创建并返回一个深度适配 PPT 生成管线的高端多模态图像生成智能体。
    """
    client = create_client(api_key)
    dashscope.api_key = client.api_key
    dashscope.base_http_api_url = client.base_url

    def generate_and_save_image(
        prompt: str, 
        *args,  # 动态吸收 main.py 传入的页码 i 和 slide.theme
        size: str = "1280*720", # 16:9 高清分辨率，避免后续插入 PPT 时发生比例拉伸
        output_dir: str = "output_images", 
        filename: str = None,
        **kwargs
    ) -> Optional[str]:
        page_index = args[0] if len(args) > 0 else None
        theme = args[1] if len(args) > 1 else None
        
        # 提取当前幻灯片专属的调色盘，注入生图提示词中
        if theme:
            try:
                # 兼容 Pydantic 对象或纯 Dict 格式
                primary_color = getattr(theme, 'primary', None) or theme.get('primary', '')
                secondary_color = getattr(theme, 'secondary', None) or theme.get('secondary', '')
                bg_color = getattr(theme, 'background', None) or theme.get('background', '')
                
                color_hints = []
                if primary_color: color_hints.append(f"主色调契合 {primary_color}")
                if secondary_color: color_hints.append(f"辅助色调呼应 {secondary_color}")
                if bg_color: color_hints.append(f"整体氛围与页面底色 {bg_color} 协调但保持清晰边界")
                
                if color_hints:
                    color_prompt = "，".join(color_hints)
                    prompt += f"，画面视觉风格需进行色彩对齐：{color_prompt}。整体呈现极简、富有呼吸感的现代视觉设计，适合作为 PPT 侧边、角落或独立插图区的美观插图，不作为背景图或文字衬底。严禁包含任何无关的嘈杂文本或文字标签。"
            except Exception as e:
                pass

        print(f"\n[Image Agent] 正在通过多模态对话协议唤醒 {model_name}...")
        if page_index is not None:
            print(f"当前服务目标：第 {page_index + 1} 页幻灯片视觉点缀图")
        print(f"最终精算提示词: {prompt}")

        messages = [
            {
                "role": "user",
                "content": [
                    {"text": prompt}
                ]
            }
        ]
        
        # 转换分辨率格式，兼容部分意外传入的 'x' 字符
        dashscope_size = size.replace("x", "*")

        try:
            response = MultiModalConversation.call(
                model=model_name,
                messages=messages,
                result_format='message',
                stream=False,
                watermark=False,
                prompt_extend=True, # 允许模型做智能意境发散扩展
                negative_prompt="低分辨率，低画质，肢体畸形，画面过饱和，无质感，低级AI感，扭曲，不合常理的排版，杂乱标签。",
                size=dashscope_size
            )

            # 6. 状态校验与响应解析
            if response.status_code == 200:
                try:
                    choices = response.output.choices
                    img_url = choices[0].message.content[0]['image']
                except (AttributeError, KeyError, IndexError) as parse_err:
                    print(f"[Image Agent] 远端握手成功但返回结构解析失败: {parse_err}")
                    print(f"[Image Agent] 原始回包: {response}")
                    return None
                
                # 本地物理缓存与离线化命名规范
                os.makedirs(output_dir, exist_ok=True)
                if not filename:
                    if page_index is not None:
                        filename = f"slide_{page_index + 1}.png"
                    else:
                        filename = f"img_{int(time.time())}.png"
                    
                file_path = os.path.join(output_dir, filename)
                
                print(f"[Image Agent] 云端渲染完成，正在拉取二进制图元数据...")
                img_response = requests.get(img_url, timeout=20)
                if img_response.status_code == 200:
                    with open(file_path, "wb") as f:
                        f.write(img_response.content)
                    
                    absolute_path = os.path.abspath(file_path)
                    print(f" [Image Agent] 图元成功离线化并写入上下文：{absolute_path}")
                    return absolute_path
                else:
                    print(f" [Image Agent] 下载物理资源失败，HTTP 状态码: {img_response.status_code}")
                    return None
            else:
                print(f" [Image Agent] 专属网关拒绝服务！HTTP返回码：{response.status_code}")
                print(f" 百炼错误码：{response.code} | 提示信息：{response.message}")
                return None
                
        except Exception as e:
            print(f"[Image Agent] 运行时管线发生未捕获异常：{e}")
            return None

    return generate_and_save_image
