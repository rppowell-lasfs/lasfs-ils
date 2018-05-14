package org.lasfs.ils.legacy.xlsimport

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Row

class RowEntry {
    def row;
    def cellItemLocation;
    def cellItemType;
    def cellItemNumber;
    def cellItemTitle;

    String itemLocation;
    String itemType;
    String itemNumber;
    String itemTitle;

    static RowEntry createRowEntry(Row row) {
        RowEntry rowEntry = new RowEntry();
        rowEntry.row = row
        rowEntry.setItemLocation(row)
        rowEntry.setItemType(row)
        rowEntry.setItemNumber(row)
        rowEntry.setItemTitle(row)

        return rowEntry;
    }


    def isValidLocation(cell) {
        return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
    }

    def getValidLocation(cell) {
        return cell.stringCellValue as String
    }

    def setItemLocation(Row row) {
        Cell cellItemLocation = row.getCell(XlsImporter.indexOfHeaderEntry("LOCATION"))
        if(isValidItemNumber(cellItemLocation)) {
            this.cellItemNumber = cellItemLocation
            itemLocation = this.getValidItemNumber(cellItemLocation)
        }
    }


    def isValidType(cell) {
        return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
    }

    def getValidType(cell) {
        return cell.stringCellValue as String
    }

    def setItemType(Row row) {
        Cell cellItemType = row.getCell(XlsImporter.indexOfHeaderEntry("TYPE"))
        if(isValidType(cellItemType)) {
            this.cellItemType = cellItemType
            itemType = this.getValidType(cellItemType)
        }
    }


    def isValidItemNumber(cell) {
        return ((cell != null) &&
                    (cell.getCellTypeEnum() == CellType.NUMERIC) &&
                    (((int) cell.numericCellValue) == cell.numericCellValue)
        )
    }

    def getValidItemNumber(cell) {
        return cell.numericCellValue as int
    }

    def setItemNumber(Row row) {
        Cell cellItemNumber = row.getCell(XlsImporter.indexOfHeaderEntry("NUMBER"))
        if(isValidItemNumber(cellItemNumber)) {
            this.cellItemNumber = cellItemNumber
            itemNumber = this.getValidItemNumber(cellItemNumber)
        }
    }


    def isValidTitle(cell) {
        return ((cell != null) && (cell.getCellTypeEnum() == CellType.STRING))
    }

    def getValidTitle(cell) {
        return cell.stringCellValue as String
    }

    def setItemTitle(Row row) {
        Cell cellItemTitle = row.getCell(XlsImporter.indexOfHeaderEntry("TITLE"))
        if(isValidTitle(cellItemTitle)) {
            this.cellItemTitle = cellItemTitle
            itemTitle = this.getValidType(cellItemTitle)
        }
    }


    ////////////////////////////////////////////////////////////////////////////////

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

}
