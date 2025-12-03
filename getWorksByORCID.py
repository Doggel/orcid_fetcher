import requests
import json
from typing import List, Dict, Optional
import pandas as pd
import sqlite3
from pathlib import Path
import time

def get_orcid_works(orcid_id: str) -> List[Dict]:
    """
    通过 ORCID Public API 获取指定 ORCID 用户的所有公开作品信息。
    
    参数:
        orcid_id (str): ORCID ID, 例如 '0000-0001-1234-5678'
    
    返回:
        List[Dict]: 每个元素包含标题, 发表时间, 期刊, DOI 等信息
    """
    # 构造 API URL
    base_url = "https://pub.orcid.org/v3.0"
    url = f"{base_url}/{orcid_id}/works"
    
    headers = {
        "Accept": "application/json"
        # "User-Agent": "APP-ABCDEFGHIJK1234 (youremail@university.edu.cn)"   # (可选) 添加 Client ID 和联系方式, 以表明应用身份，避免被限流。
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Fail to request ({orcid_id}): {e}")
        return []
    
    data = response.json()
    
    works = []
    for group in data.get("group", []):
        # Every group is a piece of work
        work_summary = group.get("work-summary", [{}])[0]

        title = work_summary.get("title", {}).get("title", {}).get("value", "N/A")

        publication_date = None
        pub_date = work_summary.get("publication-date")
        if pub_date:
            year = (pub_date.get("year") or {}).get("value") or ""
            month = (pub_date.get("month") or {}).get("value") or "01"
            day = (pub_date.get("day") or {}).get("value") or "01"
            month = str(month)
            day = str(day)
            if year:
                publication_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            else:
                publication_date = None
        
        journal_title = "N/A"
        if "journal-title" in work_summary and work_summary["journal-title"]:
            journal_title = work_summary["journal-title"].get("value", "N/A")
        else:
            source_names = work_summary.get("source", [])
            if source_names:
                source = source_names[0]
                if "source-name" in source:
                    journal_title = source["source-name"].get("value", "N/A")
        
        doi = "N/A"
        external_ids = work_summary.get("external-ids", {}).get("external-id", [])
        for ext_id in external_ids:
            if ext_id.get("external-id-type") == "doi":
                doi = ext_id.get("external-id-value", "N/A")
                break
        
        works.append({
            "title": title,
            "publication_date": publication_date,
            "journal": journal_title,
            "doi": doi
        })
    
    return works

def print_works(works: List[Dict]):
    """print works info to console"""
    for i, work in enumerate(works, 1):
        print(f"\n--- Work {i} ---")
        print(f"Title: {work['title']}")
        print(f"Data: {work['publication_date']}")
        print(f"Journal/Source: {work['journal']}")
        print(f"DOI: {work['doi']}")

def fetch_orcid_works_from_excel(input_excel: str, output_excel: str, sleep_time: float = 0.5):
    """
    从指定 Excel 文件读取 Name 和 ORCID, 获取每人作品, 并导出到新 Excel。
    
    参数:
        input_file (str): 输入 Excel 文件路径（需包含 'Name' 和 'ORCID' 列）
        output_file (str): 输出 Excel 文件路径
    """
    try:
        df_input = pd.read_excel(input_file, header=0)
        if 'Name' not in df_input.columns or 'ORCID' not in df_input.columns:
            print("❌ Excel must include 'Name' and 'ORCID' columns!")
            return
    except Exception as e:
        print(f"❌ Cannot read Excel file: {e}")
        return

    all_works = []

    for idx, row in df_input.iterrows():
        name = str(row['Name']).strip()
        orcid = str(row['ORCID']).strip()

        if not orcid or orcid.lower() in ['nan', 'none', '', 'null']:
            print(f"⚠️ Skip invalid ORCID for {name}")
            continue

        print(f"📥 Getting works for {name} ({orcid}) ...")
        works = get_orcid_works(orcid)

        for work in works:
            work.update({
                'Name': name,
                'ORCID': orcid
            })

        all_works.extend(works)
        print(f"   ➤  {len(works)} works found.")
        time.sleep(sleep_time)  # avoid rate limiting

    if all_works:
        df_output = pd.DataFrame(all_works)
        desired_order = ['Name', 'ORCID', 'title', 'publication_date', 'journal', 'doi']
        for col in desired_order:
            if col not in df_output.columns:
                df_output[col] = "N/A"
        df_output = df_output[desired_order]

        df_output.rename(columns={
            'title': 'Title',
            'publication_date': 'Publication Date',
            'journal': 'Journal/Source',
            'doi': 'DOI'
        }, inplace=True)

        df_output.to_excel(output_file, index=False)
        print(f"\n✅ Successfully exported {len(all_works)} works to {output_file}")
    else:
        print("⚠️ No works found to export.")

 
if __name__ == "__main__":
    input_file = "orcid_list.xlsx"
    output_file = "output_orcid_works.xlsx"
    fetch_orcid_works_from_excel(input_file, output_file)