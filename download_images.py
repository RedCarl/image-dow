#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
根据Excel数据批量下载图片，并按如下规则重命名保存：
文件名格式：一级-二级1-二级2-三级-品牌-条码.jpg
注意：下载前会去掉图片地址中的查询参数（?后面的内容）。
"""

import argparse
import os
import re
import sys
from typing import Dict, List, Optional

import requests
from openpyxl import load_workbook


def sanitize(text: Optional[str]) -> str:
    """对文件名中的每个字段进行清洗，去除无效字符并压缩分隔符。"""
    if text is None:
        return ""
    s = str(text).strip()
    # 替换不允许的文件名字符为破折号
    s = re.sub(r"[\\/\\:*?\"<>|]", "-", s)
    # 去除多余空格
    s = re.sub(r"\s+", "", s)
    # 压缩多个破折号
    s = re.sub(r"-+", "-", s)
    return s


def build_filename(fields: List[Optional[str]]) -> str:
    """按指定顺序拼接并返回目标文件名（固定扩展名为.jpg）。"""
    cleaned = [sanitize(f) for f in fields]
    name = "-".join(cleaned) + ".jpg"
    # 防止出现以破折号结尾或开头的异常情况
    name = re.sub(r"^-+", "", name)
    name = re.sub(r"-+\.jpg$", ".jpg", name)
    return name


def strip_query(url: str) -> str:
    """去除URL中的查询参数（?后面的内容）。"""
    if not url:
        return url
    return url.split("?")[0]


def ensure_dir(path: str) -> None:
    """确保输出目录存在。"""
    os.makedirs(path, exist_ok=True)


def find_header_indices(header_row: List[Optional[str]]) -> Dict[str, int]:
    """根据表头名称返回所需列的索引映射。"""
    header_map: Dict[str, int] = {}
    for idx, name in enumerate(header_row):
        if name is None:
            continue
        header_map[str(name).strip()] = idx
    required = ["一级", "二级1", "二级2", "三级", "品牌", "条码", "imageUrl"]
    missing = [col for col in required if col not in header_map]
    if missing:
        raise ValueError(f"Excel表头缺少必要列: {missing}")
    return header_map


def download_file(url: str, dest_path: str, timeout: int = 20) -> bool:
    """下载单个文件，成功返回True，失败返回False。"""
    try:
        resp = requests.get(url, timeout=timeout, stream=True)
        if resp.status_code != 200:
            return False
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        return True
    except Exception:
        return False


def process_excel(input_path: str, sheet_name: str, output_dir: str, start_row: int = 2, end_row: Optional[int] = None, limit: Optional[int] = None, on_progress: Optional[callable] = None, cancel_event: Optional[object] = None) -> int:
    """读取Excel并批量下载图片到指定目录。支持进度回调与取消。返回处理记录数。"""
    wb = load_workbook(input_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"工作表不存在: {sheet_name}")
    ws = wb[sheet_name]

    # 读取表头
    header_cells = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    header_idx = find_header_indices(list(header_cells))

    ensure_dir(output_dir)

    # 计算总数（含有URL的行）
    total = 0
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
        image_cell = row[header_idx["imageUrl"]]
        if image_cell not in (None, ""):
            total += 1

    processed = 0
    for row in ws.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
        if cancel_event is not None and getattr(cancel_event, "is_set", lambda: False)():
            break
        # 提取各字段
        one = row[header_idx["一级"]]
        two1 = row[header_idx["二级1"]]
        two2 = row[header_idx["二级2"]]
        three = row[header_idx["三级"]]
        brand = row[header_idx["品牌"]]
        barcode = row[header_idx["条码"]]
        image_url = strip_query(str(row[header_idx["imageUrl"]]))

        # 跳过无URL记录
        if image_url in (None, ""):
            continue

        url = image_url
        filename = build_filename([one, two1, two2, three, brand, barcode])
        dest_path = os.path.join(output_dir, filename)

        # 跳过已存在文件
        if os.path.exists(dest_path):
            if on_progress:
                on_progress({"status": "skip", "processed": processed + 1, "total": total, "filename": filename})
            else:
                print(f"已存在，跳过：{dest_path}")
            processed += 1
        else:
            ok = download_file(url, dest_path)
            if ok:
                if on_progress:
                    on_progress({"status": "success", "processed": processed + 1, "total": total, "filename": filename})
                else:
                    print(f"下载成功：{dest_path}")
                processed += 1
            else:
                if on_progress:
                    on_progress({"status": "fail", "processed": processed, "total": total, "filename": filename})
                else:
                    print(f"下载失败：{url}")

        if limit is not None and processed >= limit:
            break

    wb.close()
    if on_progress:
        on_progress({"status": "done", "processed": processed, "total": total})
    else:
        print(f"完成，处理记录数：{processed}")
    return processed


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    """解析命令行参数。"""
    parser = argparse.ArgumentParser(description="从Excel下载图片并重命名保存")
    parser.add_argument("--input", required=True, help="Excel文件路径")
    parser.add_argument("--sheet", default="Sheet1", help="工作表名称")
    parser.add_argument("--out", default="images", help="输出目录")
    parser.add_argument("--start", type=int, default=2, help="开始行（含），默认2跳过表头")
    parser.add_argument("--end", type=int, default=None, help="结束行（含），默认到最后一行")
    parser.add_argument("--limit", type=int, default=None, help="处理记录上限，仅用于试运行")
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    process_excel(
        input_path=args.input,
        sheet_name=args.sheet,
        output_dir=args.out,
        start_row=args.start,
        end_row=args.end,
        limit=args.limit,
    )


if __name__ == "__main__":
    main()