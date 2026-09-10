import os
import json
from datetime import datetime
from pydantic import BaseModel

def save_agent_log(data, filename: str, log_dir: str = "debug_logs"):
    """
    将智能体的中间输出安全地保存为本地 JSON 文件
    :param data: 可以是 Pydantic 对象、dict、list 或 字符串
    :param filename: 存储的文件名（如 '01_planner_output.json'）
    :param log_dir: 存放 debug 文件的根目录
    """
    # 确保调试目录存在
    os.makedirs(log_dir, exist_ok=True)
    filepath = os.path.join(log_dir, filename)
    
    # 智能识别数据类型并转化为标准可序列化的 dict/list
    if isinstance(data, BaseModel):
        serializable_data = data.model_dump()
    elif isinstance(data, (dict, list)):
        serializable_data = data
    else:
        # 兜底：如果是普通字符串（如报错列表），封装为字典
        serializable_data = {"raw_data": data}
        
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(serializable_data, f, ensure_ascii=False, indent=4)
        
    print(f"[Debug Log] 成功捕获并保存至: {filepath}")