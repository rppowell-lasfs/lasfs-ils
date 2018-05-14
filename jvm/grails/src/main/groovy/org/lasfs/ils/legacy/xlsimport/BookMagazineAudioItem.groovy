package org.lasfs.ils.legacy.xlsimport

import groovy.transform.Canonical

/**
 * Created by rpowell on 11/18/16.
 */

@Canonical
class BookMagazineAudioItem {
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