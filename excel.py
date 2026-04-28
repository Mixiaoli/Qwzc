import os
from datetime import datetime
import win32com.client as win32


TEMPLATE_FILE = "import_ticket_create.xlsx"


def clean_text(text):
    if not text:
        return ""
    text = str(text)
    text = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    text = " ".join(text.split())
    return text[:200]


def save_excel(data):
    if not data:
        return "⚠️ 没有数据"

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(os.path.abspath(TEMPLATE_FILE))
    ws = wb.ActiveSheet

    start_row = 2

    for i, item in enumerate(data):
        row = start_row + i

        # ==========================
        # 👉 第2条数据开始：复制模板行
        # ==========================
        if i > 0:
            ws.Rows(2).Copy()        # 复制模板行
            ws.Rows(row).Insert()    # 插入新行（带默认值）

        # ==========================
        # 清洗数据
        # ==========================
        title = clean_text(item.get("工单标题（必填）", ""))
        desc = clean_text(item.get("工单描述（必填）", ""))
        nickname = clean_text(item.get("客户昵称（昵称、邮箱或手机至少填写一个）", ""))
        email = clean_text(item.get("客户邮箱", ""))

        if not title:
            title = desc if desc else "用户问题"

        # ==========================
        # 只覆盖需要的字段（其他默认值保留）
        # ==========================
        ws.Cells(row, 1).Value = title
        ws.Cells(row, 2).Value = desc
        ws.Cells(row, 3).Value = nickname
        ws.Cells(row, 4).Value = email
        ws.Cells(row, 9).Value = item.get("受理客服", "")

    # ==========================
    # 删除多余空行
    # ==========================
    last_row = start_row + len(data)
    ws.Range(f"{last_row}:{last_row+100}").Delete()

    # ==========================
    # 保存
    # ==========================
    output_file = f"工单_{datetime.now().strftime('%H%M%S')}.xlsx"
    wb.SaveAs(os.path.abspath(output_file))

    wb.Close()
    excel.Quit()

    return f"✅ 导出成功：{output_file}"