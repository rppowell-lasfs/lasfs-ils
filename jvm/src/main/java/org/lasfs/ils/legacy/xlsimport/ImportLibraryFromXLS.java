package org.lasfs.ils.legacy.xlsimport;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;

public class ImportLibraryFromXLS {

    File xlsfile;
    Workbook workbook;

    public ImportLibraryFromXLS(File xlsfile) throws IOException {
        this.xlsfile = xlsfile;
        InputStream inputStream = new FileInputStream(xlsfile);
        workbook = WorkbookFactory.create(inputStream);

    }

    public static String CellToString(Cell c) {
        if (c.getCellType() == CellType.STRING) {
            return c.getStringCellValue();
        } else if (c.getCellType() == CellType.NUMERIC) {
            return ((new Double(c.getNumericCellValue())).toString());
        } else if (c.getCellType() == CellType.BLANK) {
            return null;
        }
        return "UNKNOWN";
    }


    public static String RowToString(Row r) {
        ArrayList<String> cellStrings = new ArrayList<String>();
        for (Cell c: r) {
            //cellStrings.add(c.getStringCellValue());
            cellStrings.add(CellToString(c));
        }
        return String.join(", ", cellStrings);
    }

    public static void main(String[] args) throws IOException {
        Path currentPath = Paths.get(System.getProperty("user.dir"));
        Path filePath = Paths.get(currentPath.toString(), "..", "legacy", "Books Magazines Audio.xls");
        System.out.println(filePath.toString());
        ImportLibraryFromXLS importer = new ImportLibraryFromXLS(filePath.toFile());

        BooksImporter booksImporter = new BooksImporter();
        Sheet booksSheet = booksImporter.getBooksSheet(importer.workbook);
        booksImporter.readBooksHeader(booksSheet);
        booksImporter.processBooksRows(booksSheet);

    }
}
