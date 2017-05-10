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

//@Field def catalog = [types: [], locations: [], ilsitems: []]
def catalog = new Catalog()

@Field def header = []
@Field def headerFlag

@Field def validItems = []
@Field def invalidItems = []

def validateHeader(sheet) {
    for (cell in sheet.getRow(0).cellIterator()) {
        header << cell.stringCellValue
    }
    println "Header:"
    println header
    headerFlag = true
}

validateHeader(sheet)

def rowProcessor = new RowProcessor(header)

for (row in sheet.rowIterator()) {
    if (headerFlag) {
        headerFlag = false
        continue
    }
    rowProcessor.processRow(row, catalog, invalidItems)
}

println "Number of items processed: " + catalog.items.size()

for (r in catalog.items[0..9]) {
    println r
}

//catalog.print_catalog()