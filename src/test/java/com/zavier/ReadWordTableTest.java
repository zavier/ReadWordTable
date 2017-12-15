package com.zavier;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Assert;
import org.junit.Test;


public class ReadWordTableTest {
    @Test
    public void test() {
        ReadWordTable readWordTable = new ReadWordTable();
        String expected =
                "<table><tr><td></td><td colspan='2'></td><td></td><td></td><td></td><td></td><td></td></tr><tr><td rowspan='25'></td><td rowspan='15'></td><td rowspan='5'></td><td></td><td rowspan='25'></td><td rowspan='25'></td><td rowspan='25'></td><td rowspan='3'></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td><td rowspan='2'></td></tr><tr><td></td></tr><tr><td rowspan='10'></td><td></td><td></td></tr><tr><td></td><td rowspan='4'></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td><td rowspan='2'></td></tr><tr><td></td></tr><tr><td></td><td rowspan='3'></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td rowspan='9'></td><td rowspan='9'></td><td></td><td></td></tr><tr><td></td><td rowspan='9'></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td></td></tr><tr><td colspan='2'></td><td></td></tr></table>";
        try (InputStream inputStream =
                ReadWordTableTest.class.getClassLoader().getResourceAsStream("table1.docx");
                XWPFDocument document = new XWPFDocument(inputStream);) {
            List<XWPFTable> tables = document.getTables();
            String tableHtml = readWordTable.readTable(tables.get(0));
            Assert.assertEquals(expected, tableHtml);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
