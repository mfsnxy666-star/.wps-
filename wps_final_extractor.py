#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WPS文件文本提取器 - 精简版
主要入口：文件流
功能：从WPS文件流中提取纯净的文本内容
"""

import os
import re
import tempfile

def extract_text_from_wps_stream(byte_stream: bytes) -> str:
    """
    从WPS文件字节流中提取文本内容
    
    Args:
        byte_stream: WPS文件的字节流数据
    
    Returns:
        提取的纯净文本内容
    """
    try:
        # 方法1: 使用OLE复合文档解析（优先方法）
        ole_text = _extract_with_ole(byte_stream)
        if ole_text and len(ole_text.strip()) > 20:
            return _clean_text(ole_text)
        
        # 方法2: 直接二进制解析（备用方法）
        binary_text = _extract_with_binary(byte_stream)
        if binary_text and len(binary_text.strip()) > 10:
            return _clean_text(binary_text)
        
        return "未能从WPS文件中提取到有效文本内容"
        
    except Exception as e:
        return f"提取失败: {str(e)}"

def _extract_with_ole(file_data: bytes) -> str:
    """使用OLE复合文档解析"""
    try:
        import olefile
        
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix='.wps') as tmp_file:
            tmp_file.write(file_data)
            tmp_file_path = tmp_file.name
        
        try:
            if not olefile.isOleFile(tmp_file_path):
                return ""
            
            ole = olefile.OleFileIO(tmp_file_path)
            text_results = []
            
            # 解析WordDocument流
            if ole.exists('WordDocument'):
                word_doc_stream = ole.openstream('WordDocument')
                stream_data = word_doc_stream.read()
                text = _parse_stream_data(stream_data)
                if text:
                    text_results.append(text)
                word_doc_stream.close()
            
            ole.close()
            return '\n'.join(text_results)
            
        finally:
            # 清理临时文件
            if os.path.exists(tmp_file_path):
                os.unlink(tmp_file_path)
        
    except ImportError:
        return ""
    except Exception:
        return ""

def _extract_with_binary(file_data: bytes) -> str:
    """直接二进制解析"""
    try:
        return _parse_stream_data(file_data)
    except Exception:
        return ""

def _parse_stream_data(data: bytes) -> str:
    """解析流数据提取文本"""
    text_segments = []
    
    # UTF-16LE解码（WPS主要编码）
    try:
        text = data.decode('utf-16le', errors='ignore')
        clean_text = _extract_meaningful_text(text)
        if clean_text:
            text_segments.append(clean_text)
    except:
        pass
    
    # UTF-8解码
    try:
        text = data.decode('utf-8', errors='ignore')
        clean_text = _extract_meaningful_text(text)
        if clean_text:
            text_segments.append(clean_text)
    except:
        pass
    
    # GBK解码
    try:
        text = data.decode('gbk', errors='ignore')
        clean_text = _extract_meaningful_text(text)
        if clean_text:
            text_segments.append(clean_text)
    except:
        pass
    
    return '\n'.join(text_segments)

def _extract_meaningful_text(text: str) -> str:
    """提取有意义的文本内容"""
    if not text or len(text.strip()) < 5:
        return ""
    
    # 分割文本为段落
    paragraphs = re.split(r'[\x00-\x1f]+', text)
    meaningful_segments = []
    
    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if len(paragraph) < 5:
            continue
        
        # 过滤格式化字符，保留中文、英文、数字、基本标点
        clean_paragraph = ''.join(c for c in paragraph 
                                if c.isprintable() or c in '\n\t ')
        
        if not clean_paragraph.strip():
            continue
        
        # 检查是否包含有意义的内容
        has_chinese = bool(re.search(r'[\u4e00-\u9fff]', clean_paragraph))
        has_english_words = bool(re.search(r'[a-zA-Z]{2,}', clean_paragraph))
        
        # 过滤元数据和格式信息
        if _is_metadata(clean_paragraph):
            continue
        
        if has_chinese or has_english_words:
            meaningful_segments.append(clean_paragraph)
    
    return '\n'.join(meaningful_segments)

def _is_metadata(text: str) -> bool:
    """判断是否为元数据或格式信息"""
    metadata_patterns = [
        r'CJOJPJQJ',
        r'Root Entry',
        r'SummaryInformation',
        r'DocumentSummaryInformation',
        r'KSOProductBuildVer',
        r'Times New Roman',
        r'Calibri',
        r'Arial',
        r'Courier New',
        r'HTML.*代码',
        r'mH.*sH.*nHtH',
        r'^\s*[0-9@#$%^&*()_+=\[\]{}|\\:";\'<>?,./`~-]+\s*$',
        r'^\s*[A-Za-z0-9@#$%^&*()_+=\[\]{}|\\:";\'<>?,./`~-]{1,3}\s*$'
    ]
    
    return any(re.search(pattern, text, re.IGNORECASE) for pattern in metadata_patterns)

def _is_garbled_text(text):
    """判断文本是否为乱码"""
    if not text or len(text) < 3:
        return True
    
    # 计算中文字符比例
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    total_chars = len(re.sub(r'\s', '', text))
    
    if total_chars == 0:
        return True
        
    chinese_ratio = chinese_chars / total_chars
    
    # 如果中文字符比例太低，可能是乱码
    if chinese_ratio < 0.4:
        return True
    
    # 检查是否包含过多的特殊符号
    special_chars = len(re.findall(r'[`\\/_<>{}|~]', text))
    if special_chars > 3:
        return True
    
    # 检查明显的乱码模式
    garbled_patterns = [
        r'[A-Za-z]{2,}\d+[A-Za-z]{2,}',  # 字母数字混合
        r'[`\\/_<>{}|~]{2,}',  # 连续特殊字符
        r'[\u4e00-\u9fff][A-Za-z]{1}[\u4e00-\u9fff][A-Za-z]{1}',  # 中英混杂模式
        r'剉|諲|鴙|蛻|颯|錘|廱|銐|恎|購|汵|蟢|蛓|筫|誰|龕|陙|馷|俌|済|烻|薡|鵞|坃|亯|擽|揯|筟|鯪|魐|鬴|昢|觺|刧'  # 明显的乱码字符
    ]
    
    for pattern in garbled_patterns:
        if re.search(pattern, text):
            return True
    
    return False

def _clean_text(text):
    """智能清理文本，保留中文文本块内的英文字母，严格过滤乱码"""
    if not text:
        return ""
    
    # 去除控制字符
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', text)
    
    # 按句号分割，逐句处理
    sentences = []
    
    # 先按句号分割
    parts = re.split(r'[。！？.!?]', cleaned)
    
    for part in parts:
        part = part.strip()
        if not part:
            continue
            
        # 只处理包含中文的部分
        if re.search(r'[\u4e00-\u9fff]', part):
            # 基础清理
            cleaned_part = re.sub(r'[^\u4e00-\u9fff\u0020-\u007e\u3000-\u303f\uff00-\uffef\s]+', ' ', part)
            cleaned_part = re.sub(r'ph333\s*', '', cleaned_part)
            cleaned_part = re.sub(r'\b[A-Za-z]{1}\b', '', cleaned_part)
            cleaned_part = re.sub(r'\s+', ' ', cleaned_part).strip()
            
            # 判断是否为乱码
            if not _is_garbled_text(cleaned_part) and len(cleaned_part) > 8:
                sentences.append(cleaned_part)
    
    # 合并句子
    result = '. '.join(sentences)
    if result and not result.endswith('.'):
        result += '.'
    
    # 最终清理
    result = re.sub(r'\s+', ' ', result)
    result = re.sub(r'[.,!?;:]{2,}', '.', result)
    
    return result.strip()

def test_wps_extraction():
    """测试WPS文件提取功能"""
    wps_file = "神经网络从零实现教程.wps"
    
    if not os.path.exists(wps_file):
        print(f"错误: 找不到文件 {wps_file}")
        return
    
    print("正在测试WPS文件文本提取...")
    print("=" * 60)
    
    # 从本地读取文件为字节流
    with open(wps_file, 'rb') as file:
        byte_data = file.read()
        text = extract_text_from_wps_stream(byte_data)
    
    print(f"✅ 提取成功! 文本长度: {len(text)} 字符")
    print("提取结果:")
    print("-" * 60)
    print(text)
    print("-" * 60)
    
    # 保存结果
    with open("wps_clean_extracted_text.txt", "w", encoding="utf-8") as f:
        f.write(text)
    print(f"\n📄 完整文本已保存到: wps_clean_extracted_text.txt")

if __name__ == "__main__":
    test_wps_extraction()