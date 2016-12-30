package org.lasfs.ils.legacy.xlsimport

import groovy.transform.Field

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

f= "../../../../../../../../" + "Books Magazines Audio.xls"
def excelFile = new File(f)

InputStream inputStream = new FileInputStream(excelFile)
Workbook wb = WorkbookFactory.create(inputStream)
Sheet sheet = wb.getSheet('Books')

@Field def header = []
@Field def headerFlag

@Field def row_values = []


def validateHeader(sheet) {
    for (cell in sheet.getRow(0).cellIterator()) {
        header << cell.stringCellValue
    }
    println "Header:"
    println header
    headerFlag = true
}

def isValidItemNumber(cell) {
    return (
            (cell != null) &&
            (cell.getCellTypeEnum() == CellType.NUMERIC) &&
            (((int)cell.numericCellValue) == cell.numericCellValue)
    )
}
def getValidItemNumber(cell) {
    return (int) cell.numericCellValue
}

def isValidLocation(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def isValidType(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def isValidTitle(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def isValidAuthor(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def isValidCoAuthor(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def isValidComments(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def processRow(row) {
    def rowData = [:]
    for (cell in row.cellIterator()) {
        def value = ''

        switch(cell.getCellTypeEnum()) {
            case CellType.STRING:
                value = cell.stringCellValue
                break
            case CellType.NUMERIC:
                value = cell.numericCellValue
                break
            default:
                value = ''
        }
        rowData << [("${header[cell.columnIndex]}".toString()): value]
        if ((cell.columnIndex == 7) && (isValidItemNumber(cell))) {
            rowData["${header[cell.columnIndex]}".toString()] = getValidItemNumber(cell)
        }
    }
    return rowData
}

validateHeader(sheet)

for (row in sheet.rowIterator()) {
    if (headerFlag) {
        headerFlag = false
        continue
    }
    row_values << processRow(row)
}

println row_values.size()

for (r in row_values[0..9]) {
    println r
}
