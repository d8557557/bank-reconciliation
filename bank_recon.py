#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""銀行匯款對帳 — 把官網「匯款訂單」CSV 整理進銀行匯款對帳總表。

用法:
    python bank_recon.py <匯款訂單.csv> [--total 對帳總表.xlsx] [--dept 電圖部]
"""
import argparse
import csv
import os
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment

SHEET_NAME = '銀行匯款對帳'
HEADERS = ['交易日', '存款金額', '摘要', '備註', '部門', '網站訂單編號', '賣家備註']
ENCODINGS = ['cp950', 'big5', 'utf-8-sig', 'utf-8']


def read_csv(path):
    last = None
    for enc in ENCODINGS:
        try:
            with open(path, encoding=enc) as f:
                return list(csv.reader(f)), enc
        except UnicodeDecodeError as e:
            last = e
    raise last


def build_index(header):
    idx = {}
    for i, h in enumerate(header):
        h = h.strip()
        if h and h not in idx:
            idx[h] = i
    return idx


def g(row, idx, name):
    i = idx.get(name)
    if i is None or i >= len(row):
        return ''
    return row[i].strip()


def clean_amount(s):
    s = (s or '').replace(',', '').strip()
    if not s:
        return None
    try:
        v = float(s)
    except ValueError:
        return s
    return int(v) if v == int(v) else v


def unwrap(s):
    if s.startswith('="') and s.endswith('"'):
        return s[2:-1]
    return s


def parse_date(s):
    d = (s or '').split(' ')[0].replace('/', '-')
    if not d:
        return ''
    y, m, day = [int(x) for x in d.split('-')]
    return datetime(y, m, day).date()


def correct_last5(note):
    """從賣家備註抓「正確後5碼」的 5 位數字，找不到回 None。"""
    if '正確' not in note:
        return None
    m = re.search(r'\d{5}', note)
    return m.group(0) if m else None


def parse_records(rows, dept):
    idx = build_index(rows[0])
    records = []
    for row in rows[1:]:
        if not any(c.strip() for c in row):
            continue
        order_no = g(row, idx, '訂單編號')
        if not order_no:
            continue
        seller_note = g(row, idx, '賣家備註')
        summary = unwrap(g(row, idx, '轉帳後五碼'))
        fixed = correct_last5(seller_note)
        if fixed:
            summary = fixed
        records.append([
            parse_date(g(row, idx, '匯款日期')),
            clean_amount(g(row, idx, '匯款金額')),
            summary,
            g(row, idx, '匯款銀行'),
            dept,
            order_no,
            seller_note,
        ])
    return records


def load_or_create(path):
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
        if SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.create_sheet(SHEET_NAME)
        if ws.max_row == 0 or ws.cell(row=1, column=1).value is None:
            ws.append(HEADERS)
        return wb, ws
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(HEADERS)
    return wb, ws


def style_rows(ws, start):
    for i in range(start, ws.max_row + 1):
        c1 = ws.cell(row=i, column=1)
        c1.alignment = Alignment(horizontal='center')
        if c1.value not in (None, ''):
            c1.number_format = 'yyyy/mm/dd'
        c2 = ws.cell(row=i, column=2)
        c2.alignment = Alignment(horizontal='right')
        if isinstance(c2.value, (int, float)):
            c2.number_format = '#,##0'
        c3 = ws.cell(row=i, column=3)
        if c3.value not in (None, ''):
            c3.value = str(c3.value)
            c3.number_format = '@'
            c3.alignment = Alignment(horizontal='center')


def main():
    ap = argparse.ArgumentParser(description='銀行匯款對帳：匯款訂單 CSV → 對帳總表')
    ap.add_argument('csv', help='官網匯出的匯款訂單 CSV')
    ap.add_argument('--total', default='銀行匯款對帳資料-總表.xlsx', help='對帳總表檔名')
    ap.add_argument('--dept', default='電圖部', help='部門固定值')
    ap.add_argument('--overwrite', action='store_true', help='清空總表既有資料列後重寫（保留表頭）')
    args = ap.parse_args()

    rows, enc = read_csv(args.csv)
    if not rows:
        print('CSV 是空的'); return 1
    records = parse_records(rows, args.dept)
    if not records:
        print('沒有可輸出的資料'); return 1

    try:
        wb, ws = load_or_create(args.total)
        if args.overwrite and ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        start = ws.max_row + 1
        for rec in records:
            ws.append(rec)
        style_rows(ws, start)
        wb.save(args.total)
    except PermissionError:
        print('寫入失敗：%s 正被 Excel 開啟，請先關閉再執行。' % args.total)
        return 1

    fixed = sum(1 for r in records if correct_last5(r[6]))
    print('來源: %s（編碼 %s）' % (os.path.basename(args.csv), enc))
    print('寫入 %d 筆 → %s（第 %d~%d 列）' % (len(records), args.total, start, ws.max_row))
    if fixed:
        print('摘要以「正確後5碼」覆蓋: %d 筆' % fixed)
    return 0


if __name__ == '__main__':
    sys.exit(main())
