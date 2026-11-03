import chardet
import pandas as pd

def detect_encoding(file_path):
    """
    自动检测CSV文件的编码，并使用正确的编码读取文件。

    参数:
    file_path (str): CSV文件的路径。

    返回:
    pandas.DataFrame: 读取成功后的DataFrame，或在失败时返回None。
    """
    # 1. 以二进制模式读取文件
    with open(file_path, 'rb') as f:
        # 只读取文件开头的一部分，通常足以判断编码
        # raw_data = f.read(10000)
        raw_data = f.read(30000)
    
    # 2. 使用 chardet 检测编码
    result = chardet.detect(raw_data)
    detected_encoding = result['encoding']
    confidence = result['confidence']
    
    # 打印检测结果以供参考
    msg = f"检测到文件[{file_path}]编码为: {detected_encoding} (置信度: {confidence:.2f})"
    
    if detected_encoding == "ascii":
        #如果检测到ascii，替换为兼容的utf-8
        detected_encoding = "utf-8"
        msg += f"检测到ascii替换为{detected_encoding}"
    
    if confidence <= 0.6:
        # 1. 以二进制模式读取文件
        with open(file_path, 'rb') as f:
            # 只读取文件开头的一部分，通常足以判断编码
            # raw_data = f.read(10000)
            raw_data = f.read(30000*10)

        # 2. 使用 chardet 检测编码
        result = chardet.detect(raw_data)
        detected_encoding = result['encoding']
        confidence = result['confidence']
        msg += f"置信度过低，重新检测,检测到文件[{file_path}]编码为: {detected_encoding} (置信度: {confidence:.2f})"
    if confidence < 0.6:
        msg += f"无法找到足够置信度的编码，请留意此信息"
        msg = "警告：" + msg
    print(msg)
    return detected_encoding,confidence