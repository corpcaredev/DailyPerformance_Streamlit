from datetime import datetime, timedelta


def is_meaningful_data(val):
    """Checks if a cell contains actual data, not just blanks, zeros, or hyphens."""
    if val is None:
        return False
    if isinstance(val, str) and val.strip() in ["", "-"]:
        return False
    if isinstance(val, (int, float)) and val == 0:
        return False
    return True


def is_date_like(v):
    if v is None: return None
    if isinstance(v, datetime): return v
    if isinstance(v, str):
        for fmt in ("%d-%b-%Y", "%d-%b-%y", "%d-%b-%Y ", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(v.strip(), fmt)
            except Exception:
                continue
    return None


def find_header_row(sheet, keyword="Scheme Name"):
    for r in range(1, 25):
        for cell in sheet[r]:
            try:
                if cell.value and str(cell.value).strip() == keyword: return r
            except Exception:
                continue
    return -1


def update_as_on_date(sheet):
    for r in range(1, 11):
        for cell in sheet[r]:
            if cell.value and str(cell.value).strip().startswith("As on"):
                yesterday = datetime.today() - timedelta(days=1)

                cell.value = f"As on {yesterday.strftime('%Y-%b-%d')}"
                return True
    return False


def get_month_id_for_column(sheet, row, col):
    for r in range(row - 1, 0, -1):
        cell = sheet.cell(row=r, column=col)
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row <= r <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                top_left_cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                if isinstance(top_left_cell.value, int) and len(str(top_left_cell.value)) == 6:
                    return top_left_cell.value
    return None




def get_parent_header_for_column(sheet, header_row, col):
    for r in range(header_row - 1, 0, -1):
        cell = sheet.cell(row=r, column=col)
        val = cell.value
        if val is None:
            for merged_range in sheet.merged_cells.ranges:
                if merged_range.min_row <= r <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                    top_left = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                    if top_left.value: return str(top_left.value).strip()
            continue
        if isinstance(val, int) and len(str(val)) == 6: continue
        if is_date_like(val): continue
        return str(val).strip()
    return None

