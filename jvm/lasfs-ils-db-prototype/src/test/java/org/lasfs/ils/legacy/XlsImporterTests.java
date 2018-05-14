package org.lasfs.ils.legacy;

import org.junit.Assert;
import org.junit.Test;
import org.lasfs.ils.legacy.xlsimport.XlsImporter;

public class XlsImporterTests {

    @Test
    public void testIndexOfHeaderEntryBooksLOCATION() {
        XlsImporter xlsImporter = new XlsImporter();
        int i = xlsImporter.indexOfHeaderEntry("LOCATION");
        Assert.assertEquals(5, i);
    }

}
