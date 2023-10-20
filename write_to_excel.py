import xlsxwriter


def write_list_of_items_to_worksheet(workbook, worksheet, items):
    bold = workbook.add_format({'bold': True})
    column_names = list(items[0].keys())
    for i, column_name in enumerate(column_names):
        worksheet.write(0, i, column_name, bold)
    for i, item in enumerate(items):
        for column_name in item.keys():
            if column_name not in column_names:
                column_names.append(column_name)
                worksheet.write(0, column_names.index(column_name), column_name, bold)
            worksheet.write(i + 1, column_names.index(column_name), item.get(column_name))


if __name__ == '__main__':
    workbook = xlsxwriter.Workbook("test.xlsx")
    worksheet = workbook.add_worksheet("test")
    items = [
        {"name": "test1", "value": "value1"},
        {"name": "test2", "bar": "value2", "value": "value2a"},
        {"name": "test3", "foo": "value3"},
        {"name": "test4", "foo": "value4"},
        {"name": "test5", "value": "value5"},
        {"name2": "test222", "value": "value222"}
    ]
    write_list_of_items_to_worksheet(workbook, worksheet, items)
    workbook.close()