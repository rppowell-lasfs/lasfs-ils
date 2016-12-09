package org.lasfs.ils.legacy.xlsimport

/**
 * Created by rpowell on 11/18/16.
 */

@Grab(group='org.apache.poi', module='poi', version='3.15')
@Grab(group='org.apache.poi', module='poi-ooxml', version='3.15')

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.CellType

println System.getProperty("user.dir");

f= "../../../../../../../" + "Books Magazines Audio - mini.xls"
def excelFile = new File(f)

InputStream inputStream = new FileInputStream(excelFile)
Workbook wb = WorkbookFactory.create(inputStream)
Sheet sheet = wb.getSheetAt(0)

def values = []
def header = []

for (cell in sheet.getRow(0).cellIterator()) {
    println cell.stringCellValue
    header << cell.stringCellValue
}

def headerFlag = true
for (row in sheet.rowIterator()) {
    if (headerFlag) {
        headerFlag = false
        continue
    }
    def rowData = [:]
    for (cell in row.cellIterator()) {
        def value = ''
        switch(cell.cellType) {
            case CellType.STRING:
                value = cell.stringCellValue
                break
            case CellType.NUMERIC:
                value = cell.numericCellValue
                break
            default:
                value = ''
        }
        rowData << ["${header[cell.columnIndex]}": value]
    }
    values << rowData
}

//Iterator<Row> rowIt = sheet.rowIterator()
//Row row = rowIt.next()
