package org.lasfs.ils.legacy;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Test;
import org.lasfs.ils.legacy.xlsimport.XlsImporter;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class ImportFromXlsTests {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        XlsImporter xlsImporter = new XlsImporter();

        System.out.println(System.getProperty("user.dir"));

        String xlsFilename = "../../" + "Books Magazines Audio.xls";

        File excelFile = new File(xlsFilename);
        InputStream inputStream = new FileInputStream(excelFile);
        Workbook workbook = WorkbookFactory.create(inputStream);

        Sheet books = xlsImporter.getBooksSheet(workbook);
        System.out.println(xlsImporter.validateBooksSheetHeader(books));

//        int count = 0;
//        for (Iterator<Row>rowIterator = books.rowIterator(); rowIterator.hasNext();) {
//            Row row = rowIterator.next();
//            if (count == 0) {
//                // header
//                continue;
//            }
//        }

        Iterator<Row>rowIterator = books.rowIterator();
        Row row;
        row = rowIterator.next(); // header
        row = rowIterator.next();

    }

}
