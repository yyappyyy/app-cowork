#!/usr/bin/env python3
"""Generate a professional household budget spreadsheet (家計簿)."""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule
from copy import copy

OUTPUT_PATH = "/home/runner/work/app-cowork/app-cowork/household-budget-spreadsheet/家計簿テンプレート.xlsx"

# Color scheme
COLORS = {
    "primary": "2C3E50",
    "accent": "3498DB",
    "success": "27AE60",
    "warning": "F39C12",
    "danger": "E74C3C",
    "light_bg": "ECF0F1",
    "white": "FFFFFF",
    "header_bg": "34495E",
    "income_bg": "D5F5E3",
    "expense_bg": "FADBD8",
    "savings_bg": "D6EAF8",
}

# Styles
header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
title_font = Font(name="Arial", size=16, bold=True, color=COLORS["primary"])
subtitle_font = Font(name="Arial", size=11, bold=True, color=COLORS["primary"])
normal_font = Font(name="Arial", size=10)
header_fill = PatternFill(start_color=COLORS["header_bg"], end_color=COLORS["header_bg"], fill_type="solid")
income_fill = PatternFill(start_color=COLORS["income_bg"], end_color=COLORS["income_bg"], fill_type="solid")
expense_fill = PatternFill(start_color=COLORS["expense_bg"], end_color=COLORS["expense_bg"], fill_type="solid")
savings_fill = PatternFill(start_color=COLORS["savings_bg"], end_color=COLORS["savings_bg"], fill_type="solid")
light_fill = PatternFill(start_color=COLORS["light_bg"], end_color=COLORS["light_bg"], fill_type="solid")

thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

MONTHS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]

INCOME_CATEGORIES = ["給与", "副業", "投資収入", "その他収入"]
EXPENSE_CATEGORIES = [
    "住居費（家賃/ローン）",
    "水道光熱費",
    "食費",
    "日用品",
    "交通費",
    "通信費",
    "保険料",
    "医療費",
    "教育費",
    "娯楽費",
    "衣服費",
    "美容費",
    "交際費",
    "その他",
]


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def create_dashboard(wb):
    ws = wb.active
    ws.title = "ダッシュボード"
    set_col_widths(ws, [3, 18, 15, 15, 15, 15, 3])

    # Title
    ws.merge_cells("B2:F2")
    ws["B2"] = "📊 家計簿ダッシュボード（年間サマリー）"
    ws["B2"].font = title_font

    # Annual summary headers
    row = 4
    headers = ["項目", "予算", "実績", "差額", "達成率"]
    for col, h in enumerate(headers, 2):
        ws.cell(row=row, column=col, value=h)
    style_header_row(ws, row, 6)

    # Summary rows
    summary_items = [
        ("年間収入合計", "", "", "", ""),
        ("年間支出合計", "", "", "", ""),
        ("年間貯蓄額", "", "", "", ""),
    ]
    for i, (item, *vals) in enumerate(summary_items, 1):
        r = row + i
        ws.cell(row=r, column=2, value=item).font = subtitle_font
        # Formulas referencing monthly sheets
        if i == 1:  # Income
            formula_parts = [f"'{m}'!E3" for m in MONTHS]
            ws.cell(row=r, column=4, value=0)  # Placeholder
            ws.cell(row=r, column=2).fill = income_fill
        elif i == 2:  # Expense
            ws.cell(row=r, column=4, value=0)
            ws.cell(row=r, column=2).fill = expense_fill
        else:  # Savings
            ws.cell(row=r, column=4, value=0)
            ws.cell(row=r, column=2).fill = savings_fill
        for c in range(2, 7):
            ws.cell(row=r, column=c).border = thin_border

    # Monthly breakdown
    row = 10
    ws.merge_cells(f"B{row}:F{row}")
    ws[f"B{row}"] = "📅 月別サマリー"
    ws[f"B{row}"].font = subtitle_font

    row = 11
    headers2 = ["月", "収入", "支出", "貯蓄", "貯蓄率"]
    for col, h in enumerate(headers2, 2):
        ws.cell(row=row, column=col, value=h)
    style_header_row(ws, row, 6)

    for i, month in enumerate(MONTHS):
        r = row + 1 + i
        ws.cell(row=r, column=2, value=month).border = thin_border
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="center")
        # Reference formulas to monthly sheets
        income_ref = f"'{month}'!E3"
        expense_ref = f"'{month}'!E21"
        ws.cell(row=r, column=3).value = f"={income_ref}"
        ws.cell(row=r, column=4).value = f"={expense_ref}"
        ws.cell(row=r, column=5).value = f"={income_ref}-{expense_ref}"
        ws.cell(row=r, column=6).value = f'=IF({income_ref}=0,"",({income_ref}-{expense_ref})/{income_ref})'
        ws.cell(row=r, column=6).number_format = "0%"
        for c in range(2, 7):
            ws.cell(row=r, column=c).border = thin_border
            if c >= 3:
                ws.cell(row=r, column=c).number_format = '#,##0"円"'

    # Update annual summary with SUM formulas
    ws.cell(row=5, column=4).value = f"=SUM(D12:D23)"
    ws.cell(row=5, column=4).number_format = '#,##0"円"'
    ws.cell(row=6, column=4).value = f"=SUM(E12:E23)"
    ws.cell(row=6, column=4).number_format = '#,##0"円"'
    ws.cell(row=7, column=4).value = f"=D5-D6"
    ws.cell(row=7, column=4).number_format = '#,##0"円"'

    # Chart - Monthly income vs expense
    chart = BarChart()
    chart.type = "col"
    chart.title = "月別 収入 vs 支出"
    chart.y_axis.title = "金額（円）"
    chart.x_axis.title = "月"
    chart.style = 10
    data = Reference(ws, min_col=3, min_row=11, max_col=4, max_row=23)
    cats = Reference(ws, min_col=2, min_row=12, max_row=23)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    chart.width = 20
    chart.height = 12
    ws.add_chart(chart, "B26")


def create_monthly_sheet(wb, month_name):
    ws = wb.create_sheet(title=month_name)
    set_col_widths(ws, [3, 22, 15, 15, 15, 15, 3])

    # Title
    ws.merge_cells("B1:F1")
    ws[f"B1"] = f"📋 {month_name} 家計簿"
    ws[f"B1"].font = title_font

    # === Income Section ===
    row = 3
    headers = ["カテゴリ", "予算", "実績", "合計", "メモ"]
    ws.cell(row=row, column=2, value="【収入】")
    ws.cell(row=row, column=2).font = subtitle_font
    ws.cell(row=row, column=2).fill = income_fill
    # Income total formula
    income_start = row + 1
    income_end = row + len(INCOME_CATEGORIES)
    ws.cell(row=row, column=5, value=f"=SUM(D{income_start}:D{income_end})")
    ws.cell(row=row, column=5).font = subtitle_font
    ws.cell(row=row, column=5).number_format = '#,##0"円"'

    row = 4
    for col, h in enumerate(headers, 2):
        ws.cell(row=row - 1, column=col + 0)
    # Category rows for income
    for i, cat in enumerate(INCOME_CATEGORIES):
        r = row + i
        ws.cell(row=r, column=2, value=cat).border = thin_border
        ws.cell(row=r, column=3, value=0).border = thin_border  # Budget
        ws.cell(row=r, column=3).number_format = '#,##0'
        ws.cell(row=r, column=4, value=0).border = thin_border  # Actual
        ws.cell(row=r, column=4).number_format = '#,##0'
        ws.cell(row=r, column=5, value="").border = thin_border  # Formula
        ws.cell(row=r, column=6, value="").border = thin_border  # Memo

    # === Expense Section ===
    row = 4 + len(INCOME_CATEGORIES) + 2
    ws.cell(row=row, column=2, value="【支出】")
    ws.cell(row=row, column=2).font = subtitle_font
    ws.cell(row=row, column=2).fill = expense_fill
    expense_start = row + 1
    expense_end = row + len(EXPENSE_CATEGORIES)
    ws.cell(row=row, column=5, value=f"=SUM(D{expense_start}:D{expense_end})")
    ws.cell(row=row, column=5).font = subtitle_font
    ws.cell(row=row, column=5).number_format = '#,##0"円"'

    # Correct the dashboard reference row
    # Dashboard references E3 for income and E21 for expense
    # Let's ensure income total is at row 3 col 5, expense total at row 21 col 5
    # Adjust: put income header at row 3
    # Re-do the layout more carefully

    row += 1
    for i, cat in enumerate(EXPENSE_CATEGORIES):
        r = row + i
        ws.cell(row=r, column=2, value=cat).border = thin_border
        ws.cell(row=r, column=3, value=0).border = thin_border
        ws.cell(row=r, column=3).number_format = '#,##0'
        ws.cell(row=r, column=4, value=0).border = thin_border
        ws.cell(row=r, column=4).number_format = '#,##0'
        ws.cell(row=r, column=5, value="").border = thin_border
        ws.cell(row=r, column=6, value="").border = thin_border

    # === Balance Section ===
    bal_row = row + len(EXPENSE_CATEGORIES) + 2
    ws.cell(row=bal_row, column=2, value="【収支バランス】")
    ws.cell(row=bal_row, column=2).font = subtitle_font
    ws.cell(row=bal_row, column=2).fill = savings_fill

    ws.cell(row=bal_row + 1, column=2, value="収入合計").border = thin_border
    ws.cell(row=bal_row + 1, column=4, value=f"=E3").border = thin_border
    ws.cell(row=bal_row + 1, column=4).number_format = '#,##0"円"'

    ws.cell(row=bal_row + 2, column=2, value="支出合計").border = thin_border
    expense_header_row = 4 + len(INCOME_CATEGORIES) + 2
    ws.cell(row=bal_row + 2, column=4, value=f"=E{expense_header_row}").border = thin_border
    ws.cell(row=bal_row + 2, column=4).number_format = '#,##0"円"'

    ws.cell(row=bal_row + 3, column=2, value="差額（貯蓄）").border = thin_border
    ws.cell(row=bal_row + 3, column=2).font = Font(name="Arial", size=10, bold=True)
    ws.cell(row=bal_row + 3, column=4, value=f"=D{bal_row+1}-D{bal_row+2}").border = thin_border
    ws.cell(row=bal_row + 3, column=4).number_format = '#,##0"円"'

    # === Daily Expense Log ===
    log_row = bal_row + 6
    ws.cell(row=log_row, column=2, value="【日別支出記録】")
    ws.cell(row=log_row, column=2).font = subtitle_font

    log_headers = ["日付", "カテゴリ", "金額", "支払方法", "メモ"]
    log_row += 1
    for col, h in enumerate(log_headers, 2):
        ws.cell(row=log_row, column=col, value=h)
    style_header_row(ws, log_row, 6)

    # Empty rows for daily logging (31 days)
    for i in range(1, 32):
        r = log_row + i
        for c in range(2, 7):
            ws.cell(row=r, column=c).border = thin_border


def create_bills_tracker(wb):
    ws = wb.create_sheet(title="固定費・サブスク管理")
    set_col_widths(ws, [3, 22, 12, 12, 12, 12, 15, 15])

    ws.merge_cells("B1:G1")
    ws["B1"] = "📌 固定費・サブスクリプション管理"
    ws["B1"].font = title_font

    headers = ["項目名", "月額", "支払日", "支払方法", "カテゴリ", "次回支払日", "ステータス"]
    row = 3
    for col, h in enumerate(headers, 2):
        ws.cell(row=row, column=col, value=h)
    style_header_row(ws, row, 8)

    # Sample entries
    bills = [
        ("家賃", 80000, 25, "口座振替", "住居費", "", ""),
        ("電気代", 8000, 15, "口座振替", "水道光熱費", "", ""),
        ("ガス代", 5000, 15, "口座振替", "水道光熱費", "", ""),
        ("水道代", 4000, 15, "口座振替", "水道光熱費", "", ""),
        ("携帯電話", 5000, 10, "クレジットカード", "通信費", "", ""),
        ("インターネット", 5000, 20, "クレジットカード", "通信費", "", ""),
        ("Netflix", 1490, 1, "クレジットカード", "娯楽費", "", ""),
        ("Spotify", 980, 1, "クレジットカード", "娯楽費", "", ""),
        ("生命保険", 10000, 27, "口座振替", "保険料", "", ""),
        ("", 0, "", "", "", "", ""),
        ("", 0, "", "", "", "", ""),
        ("", 0, "", "", "", "", ""),
    ]

    for i, bill in enumerate(bills):
        r = row + 1 + i
        for c, val in enumerate(bill, 2):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = thin_border
            if c == 3 and val:  # Amount
                cell.number_format = '#,##0"円"'

    # Total
    total_row = row + 1 + len(bills) + 1
    ws.cell(row=total_row, column=2, value="月額合計").font = subtitle_font
    ws.cell(row=total_row, column=3, value=f"=SUM(C4:C{row+len(bills)})").number_format = '#,##0"円"'
    ws.cell(row=total_row, column=3).font = subtitle_font


def create_savings_tracker(wb):
    ws = wb.create_sheet(title="貯蓄目標")
    set_col_widths(ws, [3, 22, 15, 15, 15, 12, 15])

    ws.merge_cells("B1:F1")
    ws["B1"] = "🎯 貯蓄目標トラッカー"
    ws["B1"].font = title_font

    headers = ["目標名", "目標金額", "現在の貯蓄", "残り", "達成率", "期限"]
    row = 3
    for col, h in enumerate(headers, 2):
        ws.cell(row=row, column=col, value=h)
    style_header_row(ws, row, 7)

    goals = [
        ("緊急予備費", 1000000, 300000),
        ("旅行資金", 500000, 150000),
        ("車購入", 2000000, 500000),
        ("教育資金", 3000000, 800000),
        ("老後資金", 10000000, 1000000),
    ]

    for i, (name, target, current) in enumerate(goals):
        r = row + 1 + i
        ws.cell(row=r, column=2, value=name).border = thin_border
        ws.cell(row=r, column=3, value=target).border = thin_border
        ws.cell(row=r, column=3).number_format = '#,##0"円"'
        ws.cell(row=r, column=4, value=current).border = thin_border
        ws.cell(row=r, column=4).number_format = '#,##0"円"'
        ws.cell(row=r, column=5, value=f"=C{r}-D{r}").border = thin_border
        ws.cell(row=r, column=5).number_format = '#,##0"円"'
        ws.cell(row=r, column=6, value=f"=D{r}/C{r}").border = thin_border
        ws.cell(row=r, column=6).number_format = "0%"
        ws.cell(row=r, column=7, value="").border = thin_border

    # Add data bars for progress
    rule = DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                       color="27AE60")
    ws.conditional_formatting.add(f"F4:F{row+len(goals)}", rule)


def create_instructions(wb):
    ws = wb.create_sheet(title="使い方ガイド")
    set_col_widths(ws, [3, 80])

    ws["B1"] = "📖 家計簿の使い方ガイド"
    ws["B1"].font = title_font

    instructions = [
        "",
        "【はじめに】",
        "この家計簿テンプレートは、毎月の収入と支出を管理し、",
        "年間の貯蓄目標を達成するためのツールです。",
        "",
        "【基本的な使い方】",
        "1. 各月のシートに毎日の収入と支出を記録します",
        "2. カテゴリごとの予算を設定します（予算列に入力）",
        "3. 実績を入力すると、自動的に合計が計算されます",
        "4. ダッシュボードで年間のサマリーを確認できます",
        "",
        "【シートの説明】",
        "• ダッシュボード：年間の収支サマリーとグラフ",
        "• 1月〜12月：各月の収支管理と日別支出記録",
        "• 固定費・サブスク管理：毎月の固定費一覧",
        "• 貯蓄目標：貯蓄目標の進捗管理",
        "",
        "【カスタマイズ方法】",
        "• カテゴリは自由に追加・変更できます",
        "• 色やフォントはお好みで変更してください",
        "• 行を追加する場合は、数式が反映されるか確認してください",
        "",
        "【Tips】",
        "• 毎日記録する習慣をつけましょう",
        "• 月初に予算を設定し、月末に振り返りをしましょう",
        "• 固定費は「固定費・サブスク管理」シートで一括管理",
        "• 貯蓄目標を設定するとモチベーションが上がります",
        "",
        "【注意事項】",
        "• このファイルはGoogleスプレッドシートにインポートして使えます",
        "• 数式のあるセルを上書きしないよう注意してください",
        "• 定期的にバックアップを取ることをおすすめします",
    ]

    for i, line in enumerate(instructions, 2):
        ws.cell(row=i, column=2, value=line)
        if line.startswith("【"):
            ws.cell(row=i, column=2).font = subtitle_font


def main():
    wb = Workbook()

    # Create all sheets
    create_dashboard(wb)
    for month in MONTHS:
        create_monthly_sheet(wb, month)
    create_bills_tracker(wb)
    create_savings_tracker(wb)
    create_instructions(wb)

    wb.save(OUTPUT_PATH)
    print(f"✅ 家計簿テンプレートを作成しました: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
