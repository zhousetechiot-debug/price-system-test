from __future__ import annotations

from pathlib import Path
import json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
FONT_NAME = "STSong-Light"

DEFAULT_NOTE = """1.報價支付方式
   總價低於15萬，訂金60% 尾款40%。
   總價高於15萬，訂金30% 設備進場40% 尾款30%。
2.若開立支票請開立即期票。
3.報價內容不含水電工程。(例:開關、燈具安裝)
4.自驗收發票日起算保固1年(非人為損壞)。
5.報價有效期自報價日起14日曆天止。
6.客製化設備交期為8週。
7.代購商品需全額支付，恕不退換。"""


def _local_image_path(base_dir: str, image_ref: str | None):
    if not image_ref:
        return None
    if image_ref.startswith('/uploads/'):
        p = Path(base_dir) / image_ref.lstrip('/')
        return p if p.exists() else None
    p = Path(image_ref)
    return p if p.exists() else None


def _company_logo(base_dir: str | None):
    base_dir = base_dir or ''
    for name in ['uploads/company_logo.png', 'uploads/company_logo.jpg', 'uploads/company_logo.jpeg']:
        p = Path(base_dir) / name
        if p.exists():
            return p
    return None


def build_display_items(quote):
    items = list(quote.items)
    extra = []
    try:
        curtain_rows = json.loads(getattr(quote, 'curtain_rows_json', '') or '[]')
    except Exception:
        curtain_rows = []
    for row in curtain_rows:
        qty = int(round(float(row.get('qty') or 0), 0))
        unit_price = round(float(row.get('unit_price') or 0), 0)
        line_total = round(float(row.get('line_total') or 0), 0)
        if qty and unit_price:
            space = row.get('space') or ''
            ctype = row.get('type') or '窗簾'
            track_length = int(float(row.get('track_length') or 0))
            cloth_height = int(float(row.get('cloth_height') or 0))
            pname = f"{space}-{ctype} {track_length}cm*{cloth_height}cm".strip('- ')
            extra.append(type('Obj', (), {'image_path':'','model':'','product_name':pname,'qty':qty,'unit':'組','unit_price_twd':unit_price,'line_total_twd':line_total,'note': (row.get('note') or getattr(quote, 'curtain_note', '') or '')})())
    if round(getattr(quote, 'planning_fee_total', 0) or 0):
        extra.append(type('Obj', (), {'image_path':'','model':'','product_name':'規劃費','qty':1,'unit':'式','unit_price_twd':round(getattr(quote,'planning_fee_total',0) or 0),'line_total_twd':round(getattr(quote,'planning_fee_total',0) or 0),'note':''})())
    if round(getattr(quote, 'setup_fee_total', 0) or 0):
        extra.append(type('Obj', (), {'image_path':'','model':'','product_name':'設定費','qty':1,'unit':'式','unit_price_twd':round(getattr(quote,'setup_fee_total',0) or 0),'line_total_twd':round(getattr(quote,'setup_fee_total',0) or 0),'note':''})())
    if round(getattr(quote,'dispatch_fee',0) or 0):
        extra.append(type('Obj', (), {'image_path':'','model':'','product_name':'派工費','qty':1,'unit':'式','unit_price_twd':round(getattr(quote,'dispatch_fee',0) or 0),'line_total_twd':round(getattr(quote,'dispatch_fee',0) or 0),'note':''})())
    if round(getattr(quote,'lock_install_fee',0) or 0):
        extra.append(type('Obj', (), {'image_path':'','model':'','product_name':'門鎖施工費','qty':getattr(quote,'lock_install_qty',0) or 0,'unit':'把','unit_price_twd':round(getattr(quote,'lock_install_unit_price',0) or 0),'line_total_twd':round(getattr(quote,'lock_install_fee',0) or 0),'note':''})())
    if round(getattr(quote, 'curtain_install_amount', 0) or 0) and round(getattr(quote, 'curtain_install_qty', 0) or 0):
        qty = getattr(quote, 'curtain_install_qty', 0) or 0
        unit_price = round(getattr(quote, 'curtain_install_amount', 0) or 0)
        extra.append(type('Obj', (), {'image_path':'','model':'','product_name':'窗簾施工','qty':qty,'unit':getattr(quote,'curtain_install_unit','組') or '組','unit_price_twd':unit_price,'line_total_twd':round(unit_price * qty, 0),'note':'窗簾施工'})())
    for key, label in [('weak_current','弱電費用'),('hardware','五金材料'),('water_elec','水電費用')]:
        amount = round(getattr(quote, f'{key}_amount', 0) or 0)
        if amount:
            qty = getattr(quote, f'{key}_qty', 1) or 1
            default_note = {'weak_current':'弱電材料與施工整體費用','hardware':'螺絲、固定片、耗材等','water_elec':'外部水電配合施工費用'}.get(key, '')
            extra.append(type('Obj', (), {'image_path':'','model':'','product_name':label,'qty':qty,'unit':getattr(quote,f'{key}_unit','式') or '式','unit_price_twd':round(amount/qty,0) if qty else amount,'line_total_twd':amount,'note': default_note})())
    try:
        custom_rows = json.loads(getattr(quote, 'custom_fee_json', '') or '[]')
    except Exception:
        custom_rows = []
    for row in custom_rows:
        total = round(float(row.get('total') or 0), 0)
        if total:
            extra.append(type('Obj', (), {'image_path':'','model':'','product_name':row.get('name') or '自訂工費','qty':row.get('qty') or 0,'unit':row.get('unit') or '式','unit_price_twd':round(float(row.get('unit_price') or 0),0),'line_total_twd':total,'note': row.get('note') or ''})())
    return items + extra


def export_quote_to_excel(quote, output_path: str, base_dir: str | None = None):
    wb = Workbook()
    ws = wb.active
    ws.title = "報價單"

    # A4 直接列印設定：橫式、寬度縮成 1 頁、高度自動分頁
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True

    # 匯出的 Excel 不放產品圖片，只保留適合列印的欄位
    widths = {"A": 7, "B": 15, "C": 33, "D": 9, "E": 8, "F": 12, "G": 12, "H": 34}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    thin = Side(style="thin", color="999999")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.merge_cells("C1:F1")
    ws["C1"] = "報價單"
    ws["C1"].font = Font(size=18, bold=True)
    ws["C1"].alignment = center
    logo = _company_logo(base_dir)
    if logo:
        try:
            img = XLImage(str(logo))
            img.width = 120
            img.height = 45
            ws.add_image(img, 'A1')
        except Exception:
            pass

    meta = [
        (3, '訂購單號', quote.quote_no, 'From', '智崴物聯科技股份有限公司'),
        (4, 'Name :', quote.customer_name, 'Date :', str(quote.quote_date)),
        (5, 'Contact :', quote.contact_name or '', 'ATTN :', quote.attn or ''),
        (6, 'Tel :', quote.phone or '', 'TEL :', quote.sales_phone or ''),
        (7, 'Mail :', quote.email or '', 'Mail :', quote.sales_email or ''),
        (8, 'Add :', quote.address or '', 'TEL(O) :', '(02) 2600-7081'),
        (9, '業務 :', quote.sales_name or '', '幣別', 'NTD'),
    ]
    for row, l1, v1, l2, v2 in meta:
        ws[f'A{row}'] = l1
        ws.merge_cells(f'B{row}:D{row}')
        ws[f'B{row}'] = v1
        ws[f'F{row}'] = l2
        ws.merge_cells(f'G{row}:H{row}')
        ws[f'G{row}'] = v2
        ws[f'B{row}'].alignment = left
        ws[f'G{row}'].alignment = left

    headers = ["項目", "型號", "品名", "數量", "單位", "單價", "總價", "備註"]
    start_row = 11
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = center
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row = start_row + 1
    display_items = build_display_items(quote)
    total_rows = max(15, len(display_items))
    items_per_page = 15
    for idx in range(total_rows):
        item = display_items[idx] if idx < len(display_items) else None
        ws.row_dimensions[row].height = 34
        vals = [
            idx + 1 if item else '',
            getattr(item, 'model', ''),
            getattr(item, 'product_name', ''),
            int(round(float(getattr(item, 'qty', 0) or 0), 0)) if item else '',
            getattr(item, 'unit', ''),
            round(getattr(item, 'unit_price_twd', 0)) if item else '',
            round(getattr(item, 'line_total_twd', 0)) if item else '',
            getattr(item, 'note', '') if item else ''
        ]
        note_text = str(vals[7] or '').strip()
        if note_text:
            line_count = max(2, len(note_text.splitlines()))
            ws.row_dimensions[row].height = max(34, 18 * line_count)
        for col, value in enumerate(vals, start=1):
            cell = ws.cell(row=row, column=col, value=(str(value) if col == 8 and value is not None else value))
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = center if col != 8 else left
        # 每 15 筆自動分頁；表頭用 print_title_rows 重複列印
        if idx > 0 and idx % items_per_page == 0:
            ws.row_breaks.append(Break(id=row - 1))
        row += 1

    summary = [('未稅小計', quote.subtotal), ('稅額', quote.tax_amount), ('合計', quote.total_amount)]
    if round(getattr(quote,'negotiated_total',0) or 0):
        summary.append(('議價後合計', round(getattr(quote,'negotiated_total',0) or 0)))
    if (quote.payment_scheme or '') != '一次付清':
        summary.append(('訂金', quote.deposit_1))
    if round(quote.deposit_2 or 0):
        summary.append(('設備進場', quote.deposit_2))
    if round(quote.deposit_3 or 0) and (quote.payment_scheme or '') != '一次付清':
        summary.append(('尾款(驗收款)', quote.deposit_3))
    srow = start_row + total_rows + 2
    for label, value in summary:
        ws[f'F{srow}'] = label
        ws.merge_cells(f'G{srow}:H{srow}')
        ws[f'G{srow}'] = value
        srow += 1

    note_start = srow + 1
    note_text = '備註：\n' + (quote.note or DEFAULT_NOTE)
    note_lines = note_text.splitlines()
    note_rows = max(10, len(note_lines) + 2)
    note_end = note_start + note_rows - 1

    # 備註區只合併左半邊，避免與付款資訊區重疊造成 MergedCell 錯誤
    ws.merge_cells(start_row=note_start, start_column=1, end_row=note_end, end_column=4)
    note_cell = ws.cell(row=note_start, column=1, value=note_text)
    note_cell.alignment = Alignment(wrap_text=True, vertical='top')
    note_cell.font = Font(size=10)
    for rr in range(note_start, note_end + 1):
        ws.row_dimensions[rr].height = 20

    # 付款資訊區獨立放右半邊
    payment_start = note_start
    payment_rows = [
        ('付款方式:', '匯款'),
        ('戶名:', '智崴物聯科技股份有限公司'),
        ('銀行:', '上海商業儲蓄銀行-林口分行'),
        ('帳號:', '79102-0000-66086'),
    ]
    for idx, (label, value) in enumerate(payment_rows):
        r = payment_start + idx
        ws[f'F{r}'] = label
        ws.merge_cells(f'G{r}:H{r}')
        ws[f'G{r}'] = value
        ws[f'F{r}'].alignment = left
        ws[f'G{r}'].alignment = left

    sign_row = max(note_end, payment_start + len(payment_rows) + 1) + 1
    ws[f'A{sign_row}'] = '客戶回簽'
    ws.merge_cells(f'F{sign_row}:H{sign_row}')
    ws[f'F{sign_row}'] = '(回簽後, 視同訂單生效)'
    ws[f'A{sign_row}'].alignment = left
    ws[f'F{sign_row}'].alignment = left
    ws.print_title_rows = '1:11'
    ws.print_area = f'A1:H{sign_row}'
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def export_quote_to_pdf(quote, output_path: str, base_dir: str | None = None):
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    margin = 12 * mm
    content_w = width - margin * 2

    def txt(x, y, text, size=9):
        c.setFont(FONT_NAME, size)
        c.drawString(x, y, str(text or ""))

    def draw_box(x, y, w, h):
        c.rect(x, y, w, h)

    def wrap_lines(text, max_chars):
        lines = []
        for raw in str(text or '').splitlines() or ['']:
            raw = raw.strip()
            if not raw:
                lines.append('')
                continue
            while len(raw) > max_chars:
                lines.append(raw[:max_chars])
                raw = raw[max_chars:]
            lines.append(raw)
        return lines

    logo = _company_logo(base_dir)
    if logo:
        try:
            c.drawImage(str(logo), margin, height - 28 * mm, width=42 * mm, height=14 * mm, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass
    txt(width / 2 - 12 * mm, height - 16 * mm, '報價單', 18)

    meta_top = height - 32 * mm
    row_h = 8 * mm
    left_label_w = 18 * mm
    left_value_w = 72 * mm
    right_label_w = 18 * mm
    right_value_w = 72 * mm
    x0 = margin
    x1 = x0 + left_label_w
    x2 = x1 + left_value_w
    x3 = x2 + 8 * mm
    x4 = x3 + right_label_w
    x5 = x4 + right_value_w
    meta_rows = [
        ('訂購單號', quote.quote_no, 'From', '智崴物聯科技股份有限公司'),
        ('Name :', quote.customer_name, 'Date :', str(quote.quote_date)),
        ('Contact :', quote.contact_name or '', 'ATTN :', quote.attn or ''),
        ('Tel :', quote.phone or '', 'TEL :', quote.sales_phone or ''),
        ('Mail :', quote.email or '', 'Mail :', quote.sales_email or ''),
        ('Add :', quote.address or '', 'TEL(O) :', '(02) 2600-7081'),
        ('業務 :', quote.sales_name or '', '幣別', 'NTD'),
    ]
    for idx, row in enumerate(meta_rows):
        y = meta_top - idx * row_h
        for xx, ww in [(x0,left_label_w),(x1,left_value_w),(x3,right_label_w),(x4,right_value_w)]:
            draw_box(xx, y-row_h, ww, row_h)
        txt(x0 + 2 * mm, y - 5.5 * mm, row[0], 9)
        txt(x1 + 2 * mm, y - 5.5 * mm, row[1], 9)
        txt(x3 + 2 * mm, y - 5.5 * mm, row[2], 9)
        txt(x4 + 2 * mm, y - 5.5 * mm, row[3], 9)

    table_top = meta_top - len(meta_rows) * row_h - 4 * mm
    cols = [16 * mm, 22 * mm, 26 * mm, 44 * mm, 14 * mm, 12 * mm, 18 * mm, 18 * mm, 42 * mm]
    headers = ['項目', '圖片', '型號', '品名', '數量', '單位', '單價', '總價', '備註']
    x_positions = [margin]
    for w in cols:
        x_positions.append(x_positions[-1] + w)
    total_rows = max(14, len(build_display_items(quote)))
    row_h = 11 * mm
    for i in range(total_rows + 2):
        y = table_top - i * row_h
        c.line(x_positions[0], y, x_positions[-1], y)
    for x in x_positions:
        c.line(x, table_top, x, table_top - (total_rows + 1) * row_h)
    for idx, h in enumerate(headers):
        txt(x_positions[idx] + 1.5 * mm, table_top - 7 * mm, h, 8.5)

    display_items = build_display_items(quote)
    for ridx, item in enumerate(display_items, start=1):
        y = table_top - ridx * row_h
        txt(x_positions[0] + 1.5 * mm, y - 7 * mm, ridx, 8)
        local_img = _local_image_path(base_dir or '', getattr(item, 'image_path', ''))
        if local_img:
            try:
                c.drawImage(str(local_img), x_positions[1] + 1 * mm, y - 9 * mm, width=18 * mm, height=8 * mm, preserveAspectRatio=True, mask='auto')
            except Exception:
                pass
        txt(x_positions[2] + 1 * mm, y - 7 * mm, getattr(item, 'model', ''), 8)
        for li, line in enumerate(wrap_lines(getattr(item, 'product_name', ''), 15)[:2]):
            txt(x_positions[3] + 1 * mm, y - 5 * mm - li * 3.7 * mm, line, 8)
        txt(x_positions[4] + 1 * mm, y - 7 * mm, getattr(item, 'qty', ''), 8)
        txt(x_positions[5] + 1 * mm, y - 7 * mm, getattr(item, 'unit', ''), 8)
        txt(x_positions[6] + 1 * mm, y - 7 * mm, f"{round(getattr(item, 'unit_price_twd', 0)):,}", 8)
        txt(x_positions[7] + 1 * mm, y - 7 * mm, f"{round(getattr(item, 'line_total_twd', 0)):,}", 8)
        for li, line in enumerate(wrap_lines(getattr(item, 'note', ''), 18)[:2]):
            txt(x_positions[8] + 1 * mm, y - 5 * mm - li * 3.7 * mm, line, 7)

    bottom_y = table_top - (total_rows + 1) * row_h - 3 * mm
    note_w = 118 * mm
    note_h = 64 * mm
    sum_x = margin + note_w + 6 * mm
    sum_w = content_w - note_w - 6 * mm

    draw_box(margin, bottom_y - note_h, note_w, note_h)
    txt(margin + 2 * mm, bottom_y - 5 * mm, '備註：', 9)
    for i, line in enumerate(wrap_lines(quote.note or DEFAULT_NOTE, 38)[:12]):
        txt(margin + 5 * mm, bottom_y - 10 * mm - i * 4.2 * mm, line, 7.5)

    summary_rows = [
        ('產品合計', f"{round(quote.product_subtotal or 0):,}"),
        ('規劃費', f"{round(quote.planning_fee_total or 0):,}"),
        ('設定費', f"{round(quote.setup_fee_total or 0):,}"),
        ('未稅小計', f"{round(quote.subtotal or 0):,}"),
        ('稅額', f"{round(quote.tax_amount or 0):,}"),
        ('合計', f"{round(quote.total_amount or 0):,}"),
        ('議價後合計', f"{round(getattr(quote,'negotiated_total',0) or 0):,}" if getattr(quote,'negotiated_total',0) else ''),
        ('收款方式', '30% / 40% / 30%' if quote.payment_scheme == '30/40/30' else ('60% / 40%' if str(quote.payment_scheme).startswith('60/40') else '一次付清')),
        ('訂金', f"{round(quote.deposit_1 or 0):,}"),
        ('設備進場', f"{round(quote.deposit_2 or 0):,}"),
    ]
    if quote.deposit_3:
        summary_rows.append(('尾款(驗收款)', f"{round(quote.deposit_3 or 0):,}"))
    row_h2 = 6.5 * mm
    for idx, (k, v) in enumerate(summary_rows):
        y = bottom_y - idx * row_h2
        draw_box(sum_x, y - row_h2, 26 * mm, row_h2)
        draw_box(sum_x + 26 * mm, y - row_h2, sum_w - 26 * mm, row_h2)
        txt(sum_x + 2 * mm, y - 4.5 * mm, k, 8)
        txt(sum_x + 28 * mm, y - 4.5 * mm, v, 8)

    bank_y = 18 * mm
    txt(margin, bank_y, '付款方式: 匯款', 8)
    txt(margin + 42 * mm, bank_y, '戶名: 智崴物聯科技股份有限公司', 8)
    txt(margin + 110 * mm, bank_y, '銀行: 上海商業儲蓄銀行-林口分行', 8)
    txt(margin + 170 * mm, bank_y, '帳號: 79102-0000-66086', 8)
    txt(margin, 10 * mm, '客戶回簽', 9)
    txt(width - 62 * mm, 10 * mm, '(回簽後, 視同訂單生效)', 8)
    c.save()
    return output_path

