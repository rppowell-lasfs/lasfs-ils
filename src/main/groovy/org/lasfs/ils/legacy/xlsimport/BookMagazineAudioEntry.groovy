package org.lasfs.ils.legacy.xlsimport

import groovy.transform.Canonical

/**
 * Created by rpowell on 12/9/16.
 */
@Canonical
class BookMagazineAudioEntry {
    Date outDate
    Date dueDate
    String librarian
    Date returnDate
    String itemLocation
    String itemTitle
    String itemType
    String itemNumber
    String itemAuthor
    String itemCoAuthor
    String itemComments
    String itemPublisher
    String itemSeries
    Date itemEntered
    String itemISBN
    String itemDonor
    String itemMSRP
}