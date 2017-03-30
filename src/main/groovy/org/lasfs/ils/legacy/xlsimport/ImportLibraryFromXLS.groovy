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

@Field def catalog = [types: [], locations: [], ilsitems: []]

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
    return cell.numericCellValue as int
}


def isValidLocation(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def getValidLocation(cell) {
    return cell.stringCellValue as String
}


def isValidType(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}
def getValidType(cell) {
    return cell.stringCellValue as String
}


def isValidTitle(cell) {
    return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
}

def getValidTitle(cell) {
    return cell.stringCellValue as String
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

def addLocationToCatalog(location) {
    if (!(location in catalog.locations)) {
        catalog.locations += location
    }
}
def addTypeToCatalog(type) {
    if (!(type in catalog.types)) {
        catalog.types += type
    }
}
def addItemToCatalog(item) {
    if (!(item.itemNumber in catalog.ilsitems)) {
        catalog.ilsitems[item.itemNumber] = item.itemTitle
    }
}



def processRowByCell(row) {
    def rowData = [:]

    for (cell in row.cellIterator()) {
        def value = ''

        //rowData["${header[cell.columnIndex]}".toString()] = getValidItemNumber(cell)

        if ((cell.columnIndex == 5) && (isValidLocation(cell))) {
            rowData["location"] = getValidLocation(cell)
            addLocationToCatalog(rowData["location"])
        } else if ((cell.columnIndex == 6) && (isValidType(cell))) {
            rowData["type"] = getValidType(cell)
            addTypeToCatalog(rowData["type"])
        } else if ((cell.columnIndex == 7) && (isValidItemNumber(cell))) {
            rowData["number"] = getValidItemNumber(cell)
        } else if ((cell.columnIndex == 9) && (isValidTitle(cell))) {
            rowData["title"] = getValidTitle(cell)
        } else {
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
        }
    }
    return rowData
}

def processRow(row) {
    def rowData = [:]

    cellLocation = row.getCell(5)
    cellType = row.getCell(6)
    cellItemNumber = row.getCell(7)
    cellItemTitle = row.getCell(9)

    if (isValidItemNumber(cellItemNumber) && isValidTitle(cellItemTitle)) {
        def item = [itemNumber: getValidItemNumber(cellItemNumber), itemTitle: getValidTitle(cellItemTitle)]
        //println "Adding ${item.itemNumber} - '${item.itemTitle}'"
        addItemToCatalog(item)

        if (isValidLocation(cellLocation)) {
            addLocationToCatalog(getValidLocation(cellLocation))
        }

        if (isValidType(cellType)) {
            addTypeToCatalog(getValidType(cellType))
        }
    } else {
        for (cell in row.cellIterator()) {
            def value = ''
            //rowData["${header[cell.columnIndex]}".toString()] = getValidItemNumber(cell)

            if ((cell.columnIndex == 5) && (isValidLocation(cell))) {
                rowData["location"] = getValidLocation(cell)
            } else if ((cell.columnIndex == 6) && (isValidType(cell))) {
                rowData["type"] = getValidType(cell)
            } else if ((cell.columnIndex == 7) && (isValidItemNumber(cell))) {
                rowData["number"] = getValidItemNumber(cell)
            } else if ((cell.columnIndex == 9) && (isValidTitle(cell))) {
                rowData["title"] = getValidTitle(cell)
            } else {
                switch (cell.getCellTypeEnum()) {
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
            }
        }
        println rowData
    }
    return rowData
}

def print_catalog() {
    println "catalog"
    println "Types:" + catalog.types.size()
    catalog.types.sort().each { item -> println "  '${item}'" }
    println "Locations:" + catalog.locations.size()
    catalog.locations.sort().each { item -> println "  '${item}'" }
    println "Items: " + catalog.ilsitems.size()
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

print_catalog()