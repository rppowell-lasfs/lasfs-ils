package org.lasfs.ils.legacy.xlsimport

@Grab(group = 'org.apache.poi', module = 'poi', version = '3.15')
@Grab(group = 'org.apache.poi', module = 'poi-ooxml', version = '3.15')
import org.apache.poi.ss.usermodel.CellType

/**
 * Created by rpowell on 5/5/17.



 Loan table

 OUT -
 DUE - Date
 Borrower - String
 Lib - String - librarian who updated
 Returned - Date

 | OUT | DUE | Borrower | Lib | Returned |
 |     |     |          |     |          | - item never checked out
 | Y   | Y   | Y        | Y   |          | - item is checked out
 | Y   | Y   | Y        | Y   | Y        | - item has been checked out and returned
 | Y   | Y   | Y        |     | Y        | - item has been checked out and returned - untracked librarian

 */
class RowProcessor {
    def header

    RowProcessor(theHeader) {
        header = theHeader
    }

    def isValidItemNumber(cell) {
        return (
                (cell != null) &&
                        (cell.getCellTypeEnum() == CellType.NUMERIC) &&
                        (((int) cell.numericCellValue) == cell.numericCellValue)
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

    def isValidDate(cell) {
        return((cell != null) && (cell.getCellTypeEnum() == CellType.NUMERIC))
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
        return rowData
    }

    def hasCheckoutHistory(row) {
        def cellOut = row.getCell(0)
        def cellDue = row.getCell(1)
        def cellBorrower = row.getCell(2)
        def cellLibrarian = row.getCell(3)
        def cellReturned = row.getCell(4)
        if (
            isValidDate(cellOut)
            && isValidDate(cellDue)
        //&& (cellReturned != null)
        //&& (
        //    (cellReturned.getCellTypeEnum() == CellType.BLANK)
        //    || (cellReturned.getCellTypeEnum() == CellType.STRING && cellReturned.stringCellValue == '')
        //)
        ) {
            println "loan:" + dumpRow(row)
            if (cellReturned != null) {

            }
        }
    }

    def hasCheckOutDate(row) {
        def cellOut = row.getCell(0)
        return isValidDate(cellOut)
    }
    def hasDueDate(row) {
        def cellDue = row.getCell(1)
        return isValidDate(cellDue)
    }
    def hasReturnedDate(row) {
        def cellReturned = row.getCell(4)
        return isValidDate(cellReturned)
    }

    def hasBorrower(row) {
        def cellBorrower = row.getCell(2)
        return (
            cellBorrower != null
            && (
                cellBorrower.getCellTypeEnum() == CellType.STRING
            )
        )
    }
    def hasLibrarian(row) {
        def cellLibrarian = row.getCell(3)
        return (
                cellLibrarian != null
                        && (
                        cellLibrarian.getCellTypeEnum() == CellType.STRING
                )
        )
    }


    def isCheckedOut(row) {
        def cellOut = row.getCell(0)
        def cellDue = row.getCell(1)
        def cellBorrower = row.getCell(2)
        def cellLibrarian = row.getCell(3)
        def cellReturned = row.getCell(4)
        if (
            hasCheckOutDate(row)
            //&& hasDueDate(row)
            //&& (
            //    !hasReturnedDate(row)
            //)
        ) {
            println "loan:" + [
                    index: row.getRowNum(),
                    out: hasCheckOutDate(row),
                    due: hasDueDate(row),
                    returned: hasReturnedDate(row),
                    burrower: hasBorrower(row),
                    librarian: hasLibrarian(row)
            ] + '' + dumpRow(row)
        }
    }

    def dumpRow(row) {
        def rowData = [:]
        rowData['index'] = row.getRowNum()
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
        return rowData
    }


    def extractRowData(row) {
        def rowData = [:]

        def cellLocation = row.getCell(XlsImporter.indexOfHeaderEntry("LOCATION"))
        def cellType = row.getCell(XlsImporter.indexOfHeaderEntry("TYPE"))
        def cellItemNumber = row.getCell(XlsImporter.indexOfHeaderEntry("NUMBER"))
        def cellItemTitle = row.getCell(XlsImporter.indexOfHeaderEntry("TITLE"))
    }


    def processRow(row, catalog, invalidItems) {
        def rowData = [:]

        RowEntry rowEntry = RowEntry.createRowEntry(row)
        def cellLocation = rowEntry.cellItemLocation
        def cellType = rowEntry.cellItemType
        def cellItemNumber = rowEntry.cellItemNumber
        def cellItemTitle = rowEntry.cellItemTitle

        if (isValidItemNumber(cellItemNumber) && isValidTitle(cellItemTitle)) {
            rowData = [itemNumber: getValidItemNumber(cellItemNumber), itemTitle: getValidTitle(cellItemTitle)]
            //println "Adding ${item.itemNumber} - '${item.itemTitle}'"

            if (isValidLocation(cellLocation)) {
                catalog.addLocationToCatalog(getValidLocation(cellLocation))
            }

            if (isValidType(cellType)) {
                catalog.addTypeToCatalog(getValidType(cellType))
            }

            //catalog.addItemToCatalog(rowData)
            catalog.items << rowData

            isCheckedOut(row)

/*
            if(isValidDate(cellDue)) { dateDue = cellOut.getDateCellValue() }
            if(isValidDate(cellReturned)) { dateReturned = cellOut.getDateCellValue() }

            if ((cellBorrower != null) && (cellBorrower.getCellTypeEnum() == CellType.STRING)) {
                stringBorrower = cellBorrower.stringCellValue
            }
            if ((cellLibrarian != null) && (cellLibrarian.getCellTypeEnum() == CellType.STRING)) {
                stringLibrarian = cellBorrower.stringCellValue
            }

            def loanData = [:]
*/

        } else {
            rowData = dumpRow(row)
            //println "Invalid Row: " + rowData
            invalidItems << rowData
        }
    }
}