package org.lasfs.ils.legacy.xlsimport

@Grab(group='org.apache.poi', module='poi', version='3.15')
@Grab(group='org.apache.poi', module='poi-ooxml', version='3.15')

import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.CellType


/**
 * Created by rpowell on 5/5/17.
 */
class RowProcessor {

    def cellLocation;

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

    def processRow(row, header, catalog) {
        def rowData = [:]
        def cellLocation = row.getCell(5)
        def cellType = row.getCell(6)
        def cellItemNumber = row.getCell(7)
        def cellItemTitle = row.getCell(9)

        if (isValidItemNumber(cellItemNumber) && isValidTitle(cellItemTitle)) {
            def item = [itemNumber: getValidItemNumber(cellItemNumber), itemTitle: getValidTitle(cellItemTitle)]
            //println "Adding ${item.itemNumber} - '${item.itemTitle}'"
            catalog.addItemToCatalog(item)

            if (isValidLocation(cellLocation)) {
                catalog.addLocationToCatalog(getValidLocation(cellLocation))
            }

            if (isValidType(cellType)) {
                catalog.addTypeToCatalog(getValidType(cellType))
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
}
