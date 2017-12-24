package org.acme.commercial;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;

public class Client {
    public static void main(String[] args) throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        String filepath = "D:\\poc\\test.xlsx";
        XLSBean xlsBean = new XLSBean();
        xlsBean.process(filepath);
    }
}
