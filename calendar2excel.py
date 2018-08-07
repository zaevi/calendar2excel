import calendar
import datetime
import xlsxwriter

options = {
    "first-weekday": 0,  # 每周以周几开始(0:周一 ~ 6:周日)
    "day-rows": 5,  # 每日占行数 ≥ 1
    "day-cols": 2  # 每日占列数 ≥ 1
}

styles = {
    # 头部单元格样式
    "header": {"align":"center", "bg_color":"#A6A6A6"},

    # 针对每日整个区块的样式 允许以list形式轮流变换样式
    "day-block": [{"bg_color":"#C8C8C8"},{"bg_color":"#D9D9D9"}],

    # 每日头部部分的样式, 重复项会覆盖day-block
    "day-header": {"align":"left"},   

    # 每日内容部分的样式, 重复项会覆盖day-block
    "day-text": {},

    # 非本月部分的样式
    "blank-day": {"align":"center", "valign":"vcenter", "bg_color":"#EAEAEA"}, 
}

# styles["day-block"] = [{"bg_color":"#C8C8C8"}]*5+[{"bg_color":"#D9D9D9"}]*2
# styles["blank-day"] = {"align":"center", "valign":"vcenter", "bg_color":"#EAEAEA"}

# styles["header"]["font_size"] = 12
# styles["day-header"]["font_size"] = 12
# styles["day-text"]["italic"] = True

cells_format = {}
cells_value = {}


def write_format(row, col, append_format: dict):
    cell = xlsxwriter.worksheet.xl_rowcol_to_cell_fast(row, col)
    fmt = cells_format[cell].copy() if cell in cells_format else {}
    fmt.update(append_format)
    cells_format[cell] = fmt


def write_formats(s_row, s_col, e_row, e_col, append_format: dict):
    for row in range(s_row, e_row + 1):
        for col in range(s_col, e_col + 1):
            write_format(row, col, append_format)


def write_value(row, col, value):
    cell = xlsxwriter.worksheet.xl_rowcol_to_cell_fast(row, col)
    cells_value[cell] = value


def write_finish(wb: xlsxwriter.Workbook,
                 ws: xlsxwriter.Workbook.worksheet_class):
    values, formats = set(cells_value.keys()), set(cells_format.keys())
    for c in values.difference(formats):
        ws.write(c, cells_value[c])
    for c in values.intersection(formats):
        ws.write(c, cells_value[c], wb.add_format(cells_format[c]))
    for c in formats.difference(values):
        ws.write_blank(c, None, wb.add_format(cells_format[c]))


def generate(year, month, filename):

    fwd = options["first-weekday"]
    calendar.setfirstweekday(fwd)
    weekdays = calendar.monthcalendar(year, month)

    workbook = xlsxwriter.Workbook(filename)
    ws = workbook.add_worksheet()

    rows = 1 + len(weekdays) * options["day-rows"]
    cols = 7 * options["day-cols"]

    span = options["day-cols"]
    for x in range(0, 7 * span, span):
        ws.merge_range(0, x, 0, x + span - 1, None)

    center = {"align": "center"}

    weekdays_title = ["星期" + d for d in "一二三四五六日"]
    weekdays_title = weekdays_title[fwd:] + weekdays_title[:fwd]
    for x in range(7):
        write_value(0, x * span, weekdays_title[x])
        write_format(0, x * span, styles["header"])

    for w in range(len(weekdays)):
        for d in range(7):
            y, x = 1 + w * options["day-rows"], d * span
            write_formats(y, x, y+options["day-rows"]-1, x+options["day-cols"]-1, styles["day-block"][(w*7+d)%len(styles["day-block"])])
            if weekdays[w][d] == 0:
                ws.merge_range(y, x, y+options["day-rows"]-1, x+options["day-cols"]-1, None)
                write_format(y, x, styles["blank-day"])
            else:
                write_value(y, x, str(weekdays[w][d]) + "日")
                write_format(y, x, styles["day-header"])
                for i in range(1, options["day-rows"]):
                    ws.merge_range(y+i, x, y+i, x+options["day-cols"]-1, None)
                    write_formats(y+i, x, y+i, x+options["day-cols"]-1, styles["day-text"])
            

    border_top = {"top": 1}
    border_bottom = {"bottom": 1}
    border_left = {"left": 1}
    border_right = {"right": 1}

    write_formats(0, 0, 0, cols - 1, border_top)
    write_formats(1, 0, 1, cols - 1, border_top)
    write_formats(0, 0, rows - 1, 0, border_left)

    for col in range(options["day-cols"] - 1, cols, options["day-cols"]):
        write_formats(0, col, rows - 1, col, border_right)

    for row in range(options["day-rows"], rows, options["day-rows"]):
        write_formats(row, 0, row, cols - 1, border_bottom)

    if weekdays[0][0] == 0:
        x = (weekdays[0].index(1)-1) * options["day-cols"]
        write_value(1, x, str(month)+"月")
        write_format(1,x, {"font_size":9+2*options["day-rows"]})

    write_finish(workbook, ws)

    workbook.close()


if __name__ == '__main__':
    today = datetime.date.today()
    filename = "%d-%02d.xlsx" % (today.year, today.month)
    generate(today.year, today.month, filename)
