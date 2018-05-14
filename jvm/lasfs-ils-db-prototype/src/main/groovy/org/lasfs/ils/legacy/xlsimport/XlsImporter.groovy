package org.lasfs.ils.legacy.xlsimport

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class XlsImporter {

    def sheets = ["Books"]

    static bookHeader =
            ["OUT", "DUE", "BORROWER", "LIB", "RETURNED", "LOCATION", "TYPE", "NUMBER", "G", "TITLE",
             "AUTHOR", "COAUTHOR", "Comments", "PUBLISHER", "SERIES", "ENTERED",
             "ISBN", "MSRP", "Pub. Date", "By", "Donor", "Fines", "Paid by", "???"]

    static int indexOfHeaderEntry(String s) {
        return bookHeader.indexOf(s);
    }


    def validateSheets(workbook) {

    }

    Sheet getBooksSheet(Workbook workbook) {
        Sheet sheet = workbook.getSheet("Books");
        return sheet
    }

    def validateBooksSheetHeader(sheet) {
        def header = []
        for (cell in sheet.getRow(0).cellIterator()) {
            header << cell.stringCellValue
        }
        println "Header:"
        println header
    }
}
