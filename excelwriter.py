import xlsxwriter

def writeExcel(data: list, filename: str) -> None:
    workbook = xlsxwriter.Workbook('./export/'+filename+'_excel.xlsx')
    ws = workbook.add_worksheet()

    ws.write(0, 0, 'Component')
    ws.write(0, 1, 'ARTES Parameters.2.1 Artikel nr')
    ws.write(0, 2, 'Classification Name')

    row = 1

    usedList = []

    for line in data:
        if "PropertySets" in line and "ARTES Parameters" in line['PropertySets'] and "2.1 Artikel nr" in line['PropertySets']['ARTES Parameters']:
            artNum = line['PropertySets']['ARTES Parameters']['2.1 Artikel nr']
            if not artNum in usedList:
                usedList.append(artNum)
                ws.write(row, 0, 'Any')
                ws.write(row, 1, artNum)
                ws.write(row, 2, '\'=2')
                row += 1

    workbook.close()