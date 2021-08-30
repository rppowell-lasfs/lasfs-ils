package org.lasfs.ils.legacy.xlsimport;

import java.util.Date;

public class BookLoanEntry {
    String itemNumber;
    Date outDate;
    Date dueDate;
    String borrower;
    String librarian;
    Date returnDate;
    Double fines;
    String paidBy;
}
