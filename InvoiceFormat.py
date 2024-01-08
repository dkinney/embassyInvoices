import datetime
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter
from openpyxl.styles import Border, Side, PatternFill

from InvoiceStyles import styles

thinSide = Side(style='thin', color="000000")
thickSide = Side(style='thin', color="000000")

dataStyles = {
    'CLIN': {
        'width': 10,
        'style': 'defaultCell'
    },
    'SubCLIN': {
        'width': 7,
        'style': 'defaultCell'
    },
    'Description': {
        'width': 25,
        'style': 'defaultCell'
    },
    'Category': {
        'width': 48,
        'style': 'defaultCell'
    },
    'Title': {
        'width': 28,
        'style': 'defaultCell'
    },
    'Location': {
        'width': 15,
        'style': 'defaultCell'
    },
    'City': {
        'width': 15,
        'style': 'defaultCell'
    },
    'Name': {
        'width': 18,
        'style': 'defaultCell'
    },
    'Number': {
        'width': 10,
        'style': 'defaultCell'
    },
    'Hours': {
        'width': 9,
        'style': 'numberCell'
    },
    'Percentage': {
        'width': 10,
        'style': 'percentageCell'
    },
    'Regular Rate': {
        'width': 12,
        'style': 'currencyCell'
    },
    'Overtime Rate': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Regular': {
        'width': 12,
        'style': 'numberCell'
    },
    'Regular Hours': {
        'width': 12,
        'style': 'numberCell'
    },
    'On-Call OT': {
        'width': 12,
        'style': 'numberCell'
    },
    'Scheduled OT': {
        'width': 12,
        'style': 'numberCell'
    },
    'Unscheduled OT': {
        'width': 12,
        'style': 'numberCell'
    },
    'Overtime': {
        'width': 12,
        'style': 'numberCell'
    },
    'Local Holiday': {
        'width': 12,
        'style': 'numberCell'
    },
    'Admin': {
        'width': 12,
        'style': 'numberCell'
    },
    'Subtotal': {
        'width': 12,
        'style': 'numberCell'
    },
    'Regular Wages': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Post Rate': {
        'width': 15,
        'style': 'percentageCell'
    },
    'Hazard Rate': {
        'width': 15,
        'style': 'percentageCell'
    },
    'Posting': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Hazard': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Upcharge': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Total': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Total Post': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Total Hazard': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Total Regular': {
        'width': 20,
        'style': 'currencyCell'
    },
    'Total Overtime': {
        'width': 20,
        'style': 'currencyCell'
    },
    'Total Billing': {
        'width': 20,
        'style': 'currencyCell'
    },
    'Wages': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Posting Pay': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Hazard Pay': {
        'width': 15,
        'style': 'currencyCell'
    },
    'Date': {
        'width': 15,
        'style': 'dateCell'
    },
    'ContractID': {
        'width': 20,
        'style': 'defaultCell'
    },
    'Task ID': {
        'width': 8,
        'style': 'defaultCell'
    },
    'Task Name': {
        'width': 12,
        'style': 'defaultCell'
    },
    'Descr': {
        'width': 54,
        'style': 'defaultCell'
    },
    'EmbeddedAdmin': {
        'width': 12,
        'style': 'numberCell'
    },
    'SummarySubCLIN': {
        'width': 7,
        'style': 'textCell'
    },
    'SummaryDescription': {
        'width': 25,
        'style': 'textCell'
    },
    'SummaryName': {
        'width': 18,
        'style': 'textCell'
    },
    'SummaryHours': {
        'width': 12,
        'style': 'numberCell'
    },
    'Rate': {
        'width': 10,
        'style': 'currencyCell'
    },
    'SummaryRate': {
        'width': 8,
        'style': 'currencyCell'
    },
    'SummaryTotal': {
        'width': 15,
        'style': 'currencyCell'
    },
}

def styleColumn(worksheet, column, type, rowStart = None, rowStop = None):
    start = 0 if rowStart is None else rowStart
    stop = worksheet.max_row if rowStop is None else rowStop
    style = 'defaultCell'
    width = 12

    if dataStyles.get(type) is not None:
        style = dataStyles[type]['style']
        width = dataStyles[type]['width']
    else:
        print(f'style for {type} not found')

    # print(f'styleColumn({column}, {type}, {start}, {stop})')

    for row in range(start, stop):
        worksheet[column][row].style = style

    if rowStart is None:
        worksheet.column_dimensions[column].width = width
    else:
        worksheet[column + str(rowStart + 1)].style = 'summaryTitle'

def styleRow(worksheet, row, style):
    for column in range(1, worksheet.max_column):
        worksheet.cell(row=row, column=column).style = style

def columnFunction(worksheet, column, function, amountType, dataStart = None, dataStop = None, top = False):
    if amountType == 'currency':
        style = 'currencyCellTotal'
    elif amountType == 'number':
        style = 'numberCellTotal'
    else:
        style = 'defaultCell'

    start = 1 if dataStart is None else dataStart
    stop = worksheet.max_row if dataStop is None else dataStop
    totalCell = column + str(stop + 1)

    if top:
        # the start must be at least 2
        if start == 1:
            start += 1

        totalCell = column + '1'

    worksheet[totalCell] = '=SUBTOTAL(' + function +', ' + column + str(start) + ':' + column + str(stop) + ')'
    worksheet[totalCell].style = style

def sumColumn(worksheet, column, amountType, dataStart = None, dataStop = None, top = False):
    columnFunction(worksheet, column, '109', amountType, dataStart, dataStop, top)

def formatInvoiceTab(worksheet, sheetInfo):
    worksheet.delete_cols(1, 1)

    styleColumn(worksheet, 'A', 'SubCLIN')
    styleColumn(worksheet, 'B', 'Description')
    styleColumn(worksheet, 'C', 'Name')
    styleColumn(worksheet, 'D', 'Hours')
    styleColumn(worksheet, 'E', 'Rate')
    styleColumn(worksheet, 'F', 'Total')

    rowsToSum = sheetInfo['rowsToSum']

    # add SUM() formulas
    for row in rowsToSum:
        sumColumn(worksheet, 'D', 'number', row[0], row[1])
        sumColumn(worksheet, 'F', 'currency', row[0], row[1])
        
        for column in ['A', 'B', 'C', 'D', 'E', 'F']:
            for r in range(row[0]-1, row[1]):
                worksheet[column + str(r + 1)].border = Border(left=thinSide, top=thinSide, right=thinSide, bottom=thinSide)
    
    worksheet[f'B{worksheet.max_row}'].style = 'invoiceSummaryText'
    worksheet[f'D{worksheet.max_row}'].style = 'invoiceSummaryNumber'
    worksheet[f'F{worksheet.max_row}'].style = 'invoiceSummaryCurrency'

    logo = openpyxl.drawing.image.Image('logo-MEC.png')
    worksheet.add_image(logo, 'A1')

    worksheet['F1'] = 'Invoice'
    worksheet['F1'].style = 'invoiceTitle'

    worksheet['D3'] = 'Invoice Date:'
    worksheet['D3'].style = 'invoiceHeader'
    worksheet['E3'] = datetime.datetime.now().strftime("%m/%d/%Y")
    worksheet['E3'].style = 'invoiceValue'
    worksheet['D4'] = 'Invoice Number:'
    worksheet['D4'].style = 'invoiceHeader'
    worksheet['E4'] = sheetInfo['invoiceNumber']
    worksheet['E4'].style = 'invoiceValue'
    worksheet['D5'] = 'Invoice Amount:'
    worksheet['D5'].style = 'invoiceHeader'
    worksheet.merge_cells('E5:F5')
    worksheet['E5'] = sheetInfo['invoiceAmount']
    worksheet['E5'].style = 'invoiceAmount'
    worksheet['D6'] = 'Contract Number:'
    worksheet['D6'].style = 'invoiceHeader'
    worksheet['E6'] = '19AQMM23C00417'
    worksheet['E6'].style = 'invoiceValue'
    worksheet['D7'] = 'Task Order:'
    worksheet['D7'].style = 'invoiceHeader'
    worksheet['E7'] = sheetInfo['taskOrder']
    worksheet['E7'].style = 'invoiceValue'
    worksheet['D8'] = 'Billing From:'
    worksheet['D8'].style = 'invoiceHeader'
    worksheet['E8'] = sheetInfo['dateStart'] + ' - ' + sheetInfo['dateEnd']
    worksheet['E8'].style = 'invoiceValue'
    worksheet['D9'] = 'Payment Terms:'
    worksheet['D9'].style = 'invoiceHeader'
    worksheet['E9'] = 'Net 30'
    worksheet['E9'].style = 'invoiceValue'

    row = 3
    worksheet['B' + str(row)] = 'MEC Energy Services'; row += 1
    worksheet['B' + str(row)] = '3949 Hwy 8, Suite 110'; row += 1
    worksheet['B' + str(row)] = 'New Town, ND 58763'; row += 1
    worksheet['B' + str(row)] = 'TIN: 753209819'; row += 1

    row += 4; toRow = row
    worksheet['A' + str(row)] = 'Bill To:'
    worksheet['A' + str(row)].style = 'invoiceHeader'
    worksheet['B' + str(row)] = 'IPP'; row += 1
    worksheet['B' + str(row)] = 'Global Financial Services Center'; row += 1
    worksheet['B' + str(row)] = 'P.O. Box 150008'; row += 1
    worksheet['B' + str(row)] = 'ATTN: Office of Claims'; row += 1
    worksheet['B' + str(row)] = 'Charleston, SC 29415-5008'; row += 1
    worksheet['B' + str(row)] = 'Re: Helga Lumpkin'; row += 1

    row += 1; instructionsRow = row
    worksheet['A' + str(row)] = 'ACH:'
    worksheet['A' + str(row)].style = 'invoiceHeader'
    worksheet['B' + str(row)] = 'Wells Fargo'; row += 1
    worksheet['A' + str(row)].style = 'invoiceHeader'
    worksheet['B' + str(row)] = 'Routing #: 121000248'; row += 1
    worksheet['B' + str(row)] = 'Account #: 299912421028'

    row = toRow
    worksheet['D' + str(row)] = 'Remit To:'
    worksheet['D' + str(row)].style = 'invoiceHeader'
    worksheet['E' + str(row)] = 'Accounts Receivable'; row += 1
    worksheet['E' + str(row)] = 'MEC Energy Services'; row += 1
    worksheet['E' + str(row)] = '23808 Andrew Road, Unit 3'; row += 1
    worksheet['E' + str(row)] = 'Plainfield, IL 60585'; row += 1

    row = instructionsRow
    worksheet['D' + str(row)] = 'Invoice Questions:'
    worksheet['D' + str(row)].style = 'invoiceHeader'
    worksheet['E' + str(row)] = 'Joe Santorelli'; row += 1
    worksheet['E' + str(row)] = 'joe.santorelli@mandaree.com'; row += 1
    worksheet['E' + str(row)] = '478.714.0070'; row += 1

    row += 1
    worksheet['A' + str(row)] = 'CLIN'
    worksheet['A' + str(row)].style = 'summaryTitle'
    worksheet['B' + str(row)] = 'Category'
    worksheet['B' + str(row)].style = 'summaryTitle'
    worksheet['C' + str(row)] = 'Name'
    worksheet['C' + str(row)].style = 'summaryTitle'
    worksheet['D' + str(row)] = 'Hours'
    worksheet['D' + str(row)].style = 'summaryTitle'
    worksheet['E' + str(row)] = 'Rate'
    worksheet['E' + str(row)].style = 'summaryTitle'
    worksheet['F' + str(row)] = 'Amount'
    worksheet['F' + str(row)].style = 'summaryTitle'

def formatCostsTab(worksheet, sheetInfo):
    worksheet.delete_cols(1, 1)

    styleColumn(worksheet, 'A', 'SubCLIN')
    styleColumn(worksheet, 'B', 'Description')
    styleColumn(worksheet, 'C', 'SubCLIN')
    styleColumn(worksheet, 'D', 'Total')
    styleColumn(worksheet, 'E', 'Total')
    styleColumn(worksheet, 'F', 'Total')

    rowsToSum = sheetInfo['rowsToSum']

    # add SUM() formulas
    for row in rowsToSum:
        sumColumn(worksheet, 'D', 'currency', row[0], row[1])
        sumColumn(worksheet, 'E', 'currency', row[0], row[1])
        sumColumn(worksheet, 'F', 'currency', row[0], row[1])
        
        for column in ['A', 'B', 'C', 'D', 'E', 'F']:
            for r in range(row[0]-1, row[1]):
                worksheet[column + str(r + 1)].border = Border(left=thinSide, top=thinSide, right=thinSide, bottom=thinSide)
    
    worksheet[f'B{worksheet.max_row}'].style = 'invoiceSummaryText'
    worksheet[f'D{worksheet.max_row}'].style = 'invoiceSummaryCurrency'
    worksheet[f'E{worksheet.max_row}'].style = 'invoiceSummaryCurrency'
    worksheet[f'F{worksheet.max_row}'].style = 'invoiceSummaryCurrency'

    logo = openpyxl.drawing.image.Image('logo-MEC.png')
    worksheet.add_image(logo, 'A1')

    worksheet['F1'] = 'Invoice'
    worksheet['F1'].style = 'invoiceTitle'

    worksheet['D3'] = 'Invoice Date:'
    worksheet['D3'].style = 'invoiceHeader'
    worksheet['E3'] = datetime.datetime.now().strftime("%m/%d/%Y")
    worksheet['E3'].style = 'invoiceValue'
    worksheet['D4'] = 'Invoice Number:'
    worksheet['D4'].style = 'invoiceHeader'
    worksheet['E4'] = sheetInfo['invoiceNumber']
    worksheet['E4'].style = 'invoiceValue'
    worksheet['D5'] = 'Invoice Amount:'
    worksheet['D5'].style = 'invoiceHeader'
    worksheet.merge_cells('E5:F5')
    worksheet['E5'] = sheetInfo['invoiceAmount']
    worksheet['E5'].style = 'invoiceAmount'
    worksheet['D6'] = 'Contract Number:'
    worksheet['D6'].style = 'invoiceHeader'
    worksheet['E6'] = '19AQMM23C00417'
    worksheet['E6'].style = 'invoiceValue'
    worksheet['D7'] = 'Task Order:'
    worksheet['D7'].style = 'invoiceHeader'
    worksheet['E7'] = sheetInfo['taskOrder']
    worksheet['E7'].style = 'invoiceValue'
    worksheet['D8'] = 'Billing From:'
    worksheet['D8'].style = 'invoiceHeader'
    worksheet['E8'] = sheetInfo['dateStart'] + ' - ' + sheetInfo['dateEnd']
    worksheet['E8'].style = 'invoiceValue'
    worksheet['D9'] = 'Payment Terms:'
    worksheet['D9'].style = 'invoiceHeader'
    worksheet['E9'] = 'Net 30'
    worksheet['E9'].style = 'invoiceValue'

    row = 3
    worksheet['B' + str(row)] = 'MEC Energy Services'; row += 1
    worksheet['B' + str(row)] = '3949 Hwy 8, Suite 110'; row += 1
    worksheet['B' + str(row)] = 'New Town, ND 58763'; row += 1
    worksheet['B' + str(row)] = 'TIN: 753209819'; row += 1

    row += 4; toRow = row
    worksheet['A' + str(row)] = 'Bill To:'
    worksheet['A' + str(row)].style = 'invoiceHeader'
    worksheet['B' + str(row)] = 'IPP'; row += 1
    worksheet['B' + str(row)] = 'Global Financial Services Center'; row += 1
    worksheet['B' + str(row)] = 'P.O. Box 150008'; row += 1
    worksheet['B' + str(row)] = 'ATTN: Office of Claims'; row += 1
    worksheet['B' + str(row)] = 'Charleston, SC 29415-5008'; row += 1
    worksheet['B' + str(row)] = 'Re: Helga Lumpkin'; row += 1

    row += 1; instructionsRow = row
    worksheet['A' + str(row)] = 'ACH:'
    worksheet['A' + str(row)].style = 'invoiceHeader'
    worksheet['B' + str(row)] = 'Wells Fargo'; row += 1
    worksheet['A' + str(row)].style = 'invoiceHeader'
    worksheet['B' + str(row)] = 'Routing #: 121000248'; row += 1
    worksheet['B' + str(row)] = 'Account #: 299912421028'

    row = toRow
    worksheet['D' + str(row)] = 'Remit To:'
    worksheet['D' + str(row)].style = 'invoiceHeader'
    worksheet['E' + str(row)] = 'Accounts Receivable'; row += 1
    worksheet['E' + str(row)] = 'MEC Energy Services'; row += 1
    worksheet['E' + str(row)] = '23808 Andrew Road, Unit 3'; row += 1
    worksheet['E' + str(row)] = 'Plainfield, IL 60585'; row += 1

    row = instructionsRow
    worksheet['D' + str(row)] = 'Invoice Questions:'
    worksheet['D' + str(row)].style = 'invoiceHeader'
    worksheet['E' + str(row)] = 'Joe Santorelli'; row += 1
    worksheet['E' + str(row)] = 'joe.santorelli@mandaree.com'; row += 1
    worksheet['E' + str(row)] = '478.714.0070'; row += 1

    row += 1
    worksheet['A' + str(row)] = 'CLIN'
    worksheet['A' + str(row)].style = 'summaryTitle'
    worksheet['B' + str(row)] = 'Location'
    worksheet['B' + str(row)].style = 'summaryTitle'
    worksheet['C' + str(row)] = 'Type'
    worksheet['C' + str(row)].style = 'summaryTitle'
    worksheet['D' + str(row)] = 'Amount'
    worksheet['D' + str(row)].style = 'summaryTitle'
    worksheet['E' + str(row)] = 'G&A'
    worksheet['E' + str(row)].style = 'summaryTitle'
    worksheet['F' + str(row)] = 'Total'
    worksheet['F' + str(row)].style = 'summaryTitle'

def formatDetailTab(worksheet):
    worksheet.delete_cols(1, 1)
    worksheet.insert_rows(1, 1)

    styleColumn(worksheet, 'A', 'Date')
    styleColumn(worksheet, 'B', 'CLIN')
    styleColumn(worksheet, 'C', 'City')
    styleColumn(worksheet, 'D', 'SubCLIN')
    styleColumn(worksheet, 'E', 'Category')
    styleColumn(worksheet, 'F', 'Name')

    styleColumn(worksheet, 'G', 'Hours')
    styleColumn(worksheet, 'H', 'Hours')
    styleColumn(worksheet, 'I', 'Hours')
    styleColumn(worksheet, 'J', 'Hours')
    styleColumn(worksheet, 'K', 'Hours')
    styleColumn(worksheet, 'L', 'Hours')

    styleColumn(worksheet, 'M', 'Hours')
    styleColumn(worksheet, 'N', 'Hours')
    styleColumn(worksheet, 'O', 'Hours')
    styleColumn(worksheet, 'P', 'Hours')

    styleColumn(worksheet, 'Q', 'Hours')
    styleColumn(worksheet, 'R', 'Hours')
    styleColumn(worksheet, 'S', 'Hours')
    
    styleColumn(worksheet, 'T', 'Rate')
    styleColumn(worksheet, 'U', 'Total')
    styleColumn(worksheet, 'V', 'Total')

    # styleColumn(worksheet, 'W', 'Hours')
    # styleColumn(worksheet, 'X', 'Total')

    # styleColumn(worksheet, 'Y', 'Rate')
    # styleColumn(worksheet, 'Z', 'Percentage')
    # styleColumn(worksheet, 'AA', 'Total')
    # styleColumn(worksheet, 'AB', 'Percentage')
    # styleColumn(worksheet, 'AC', 'Total')

    # create a table
    table = Table(displayName='Detail', ref="A2:" + get_column_letter(worksheet.max_column) + str(worksheet.max_row))
    worksheet.add_table(table)
    worksheet.freeze_panes = worksheet['A3']

    # add SUM() formulas
    start = 3
    stop = worksheet.max_row
    sumColumn(worksheet, 'G', 'number', start, stop, top=True)
    sumColumn(worksheet, 'H', 'number', start, stop, top=True)
    sumColumn(worksheet, 'I', 'number', start, stop, top=True)
    sumColumn(worksheet, 'J', 'number', start, stop, top=True)
    sumColumn(worksheet, 'K', 'number', start, stop, top=True)
    sumColumn(worksheet, 'L', 'number', start, stop, top=True)
    sumColumn(worksheet, 'M', 'number', start, stop, top=True)
    sumColumn(worksheet, 'N', 'number', start, stop, top=True)
    sumColumn(worksheet, 'O', 'number', start, stop, top=True)
    sumColumn(worksheet, 'P', 'number', start, stop, top=True)
    sumColumn(worksheet, 'Q', 'number', start, stop, top=True)
    sumColumn(worksheet, 'R', 'number', start, stop, top=True)
    sumColumn(worksheet, 'S', 'number', start, stop, top=True)
    sumColumn(worksheet, 'U', 'currency', start, stop, top=True)
    sumColumn(worksheet, 'V', 'currency', start, stop, top=True)
    # sumColumn(worksheet, 'W', 'number', start, stop, top=True)
    # sumColumn(worksheet, 'X', 'currency', start, stop, top=True)
    # sumColumn(worksheet, 'AA', 'currency', start, stop, top=True)
    # sumColumn(worksheet, 'AC', 'currency', start, stop, top=True)