package org.lasfs.ils.legacy.xlsimport;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

public class BooksImporter {
    public void readBooksHeader(Sheet sheet) {
        ArrayList<String> headerCellStrings = new ArrayList<String>();
        for (Cell c: sheet.getRow(0)) {
            c.getStringCellValue();
            headerCellStrings.add(c.getStringCellValue());
        }
        System.out.println(headerCellStrings);
    }

    public boolean processBooksEntryNumber(Row row, BookEntry b) {
        Cell c = row.getCell(7);
        if (c.getCellType() == CellType.NUMERIC) {
            double d = c.getNumericCellValue();
            if ((d == Math.floor(d)) && !Double.isInfinite(d)) {
                b.itemNumber = (int)d;
                return true;
            }
        }
        return false;
    }

    public boolean processBooksEntryTitle(Row row, BookEntry b) {
        Cell c = row.getCell(9);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemTitle = c.getStringCellValue();
            return true;
        }
        return false;
    }

    public boolean processBooksEntryAuthor(Row row, BookEntry b) {
        Cell c = row.getCell(10);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemAuthor = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksEntryCoauthor(Row row, BookEntry b) {
        Cell c = row.getCell(11);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemCoauthor = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksEntryPublisher(Row row, BookEntry b) {
        Cell c = row.getCell(13);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemPublisher = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksEntryLocation(Row row, BookEntry b) {
        Cell c = row.getCell(5);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemLocation = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksEntryType(Row row, BookEntry b) {
        Cell c = row.getCell(6);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemType = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksEntryComments(Row row, BookEntry b) {
        Cell c = row.getCell(14);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.itemComments = c.getStringCellValue();
            return true;
        }
        return false;
    }

    public boolean processBooksLoanEntryOut(Row row, BookLoanEntry b) {
        Cell c = row.getCell(0);
        if (c.getCellType() == CellType.FORMULA) {
            FormulaEvaluator e = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellType ct = e.evaluateFormulaCell(c);
            if (ct == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(c)) {
                    b.outDate = c.getDateCellValue();
                    return true;
                }
            }
        }
        if (c != null && c.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(c)) {
                b.outDate = c.getDateCellValue();
                return true;
            }
        } else if (c != null && c.getCellType() == CellType.STRING) {
            System.out.println(c.getStringCellValue());
            return false;
        }
        return false;
    }

    public boolean processBooksLoanEntryDue(Row row, BookLoanEntry b) {
        Cell c = row.getCell(1);
        if (c.getCellType() == CellType.FORMULA) {
            FormulaEvaluator e = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellType ct = e.evaluateFormulaCell(c);
            if (ct == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(c)) {
                    b.dueDate = c.getDateCellValue();
                    return true;
                }
            }
        }
        if (c != null && c.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(c)) {
                b.dueDate = c.getDateCellValue();
                return true;
            }
        } else if (c != null && c.getCellType() == CellType.STRING) {
            System.out.println(c.getStringCellValue());
            return false;
        }
        return false;
    }
    public boolean processBooksLoanEntryBorrower(Row row, BookLoanEntry b) {
        Cell c = row.getCell(2);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.borrower = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksLoanEntryLibrarian(Row row, BookLoanEntry b) {
        Cell c = row.getCell(3);
        if (c != null && c.getCellType() == CellType.STRING) {
            b.librarian = c.getStringCellValue();
            return true;
        }
        return false;
    }
    public boolean processBooksLoanEntryReturned(Row row, BookLoanEntry b) {
        Cell c = row.getCell(4);
        if (c != null && c.getCellType() == CellType.FORMULA) {
            FormulaEvaluator e = row.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellType ct = e.evaluateFormulaCell(c);
            if (ct == CellType.NUMERIC) {
                if (DateUtil.isCellDateFormatted(c)) {
                    b.returnDate = c.getDateCellValue();
                    return true;
                }
            }
        }
        if (c != null && c.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(c)) {
                b.returnDate = c.getDateCellValue();
                return true;
            }
        } else if (c != null && c.getCellType() == CellType.STRING) {
            System.out.println(c.getStringCellValue());
            return false;
        }
        return false;
    }

    public void processBooksRow(Sheet sheet, int i) {
        Row r = sheet.getRow(i);

        BookEntry entry = new BookEntry();
        BookLoanEntry loanEntry = new BookLoanEntry();

//        System.out.println("" + i + " " + ImportLibraryFromXLS.RowToString(r));
        processBooksEntryNumber(r, entry);
        processBooksEntryTitle(r, entry);

        processBooksEntryAuthor(r, entry);
        processBooksEntryCoauthor(r, entry);
        processBooksEntryPublisher(r, entry);

        processBooksEntryLocation(r, entry);
        processBooksEntryType(r, entry);
        processBooksEntryComments(r, entry);

        processBooksLoanEntryOut(r, loanEntry);
        processBooksLoanEntryDue(r, loanEntry);
        processBooksLoanEntryBorrower(r, loanEntry);
        processBooksLoanEntryLibrarian(r, loanEntry);
        processBooksLoanEntryReturned(r, loanEntry);

        System.out.println(
                new ArrayList<String>( Arrays.asList(
                        String.valueOf(i)
                )) + " " +
                new ArrayList<String>( Arrays.asList(
                        String.valueOf(entry.itemNumber), entry.itemTitle,
                        entry.itemAuthor, entry.itemCoauthor, entry.itemPublisher,
                        entry.itemLocation, entry.itemType, entry.itemComments
                )) + " " +
                new ArrayList<String>( Arrays.asList(
                        String.valueOf(loanEntry.outDate), String.valueOf(loanEntry.dueDate),
                        loanEntry.borrower, loanEntry.librarian, String.valueOf(loanEntry.returnDate)
                ))
        );

    }


    public void processBooksRows(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        System.out.println("sheet.getLastRowNum(" + lastRowNum + ")");

        processBooksRow(sheet, 6897-1);
//        processBooksRow(sheet, lastRowNum);
//        for (int i = 1; i <= lastRowNum; i++) {
//            processBooksRow(sheet, i);
//        }

    }

    public Sheet getBooksSheet(Workbook workbook) {
        FormulaEvaluator evaluator=workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        Sheet sheet = workbook.getSheet("Books");
        return sheet;
    }
}
