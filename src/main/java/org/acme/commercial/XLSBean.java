package org.acme.commercial;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

@Slf4j
public class XLSBean {
    public void process(String filePath) throws OpenXML4JException, IOException, SAXException, ParserConfigurationException {
        try (OPCPackage p = OPCPackage.open(new File(filePath), PackageAccess.READ)) {
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(p);
            XSSFReader xssfReader = new XSSFReader(p);
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int index = 0;
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    DataFormatter formatter = new DataFormatter();
                    InputSource sheetSource = new InputSource(stream);
                    XMLReader sheetParser = SAXHelper.newXMLReader();
                    ContentHandler handler = new XSSFSheetXMLHandler(styles, null, strings, new XLSContentHandler(), formatter, false);
                    sheetParser.setContentHandler(handler);
                    sheetParser.parse(sheetSource);
                }
                ++index;
            }
        }
    }
}
