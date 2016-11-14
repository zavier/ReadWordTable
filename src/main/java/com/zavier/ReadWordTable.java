package com.zavier;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;

/**
 * 读取word中的表格，包括复杂表格（合并的单元格）
 */
public class ReadWordTable {

    /**
     * 保存需要被忽略的单元格
     */
    private List<String> omitCellsList = new ArrayList<>();

    /**
     * 表格行数
     */
    private int tableRows = 0;

    /**
     * 生成忽略的单元格列表中的格式
     * 
     * @param row
     * @param col
     * @return
     */
    public String generateOmitCellStr(int row, int col) {
        return row + ":" + col;
    }

    /**
     * 获取当前单元格的colspan属性
     * 
     * @param tcPr
     * @return
     */
    public int getColspan(CTTcPr tcPr) {
        // 判断是否存在列合并
        CTDecimalNumber gridSpan = null;
        if ((gridSpan = tcPr.getGridSpan()) != null) { // 合并的起始列
            // 获取合并的列数
            BigInteger num = gridSpan.getVal();
            return num.intValue();
        } else { // 其他被合并的列或正常列
            return 1;
        }
    }

    /**
     * 获取当前单元格的rowspan属性
     * 
     * @param table
     * @param row
     * @param col
     * @param list
     */
    public void getRowspan(XWPFTable table, int row, int col, List<Boolean> list) {
        if (++row >= tableRows) {
            return;
        }
        XWPFTableCell tableCell = table.getRow(row).getCell(col);

        // 如果此列已经被合并，则得到null
        if (tableCell == null) {
            return;
        }

        CTTcPr tcPr = tableCell.getCTTc().getTcPr();
        CTVMerge vMerge = tcPr.getVMerge();
        if (vMerge != null) { // 下一行可能是起始合并行或者中间合并行
            if (vMerge.getVal() == null) { // 是中间行，rowspan加1
                list.add(true);
                getRowspan(table, row, col, list);
            }
        }
    }


    /**
     * 添加忽略的单元格(poi区分不出合并的行，故需要手动区分)
     * 
     * @param row
     * @param col
     * @param rowspan
     */
    public void addOmitCell(int row, int col, int rowspan) {
        for (int i = row + 1; i < row + rowspan; i++) {
            String omitCellStr = generateOmitCellStr(i, col);
            omitCellsList.add(omitCellStr);
        }
    }

    public boolean isOmitCell(int row, int col) {
        String cellStr = generateOmitCellStr(row, col);
        return omitCellsList.contains(cellStr);
    }

    public String readTable(XWPFTable table) throws IOException {
        // 表格行数
        int tableRowsSize = table.getRows().size();
        tableRows = tableRowsSize;
        StringBuilder tableToHtmlStr = new StringBuilder("<table>");

        for (int i = 0; i < tableRowsSize; i++) {
            tableToHtmlStr.append("<tr>");
            int tableCellsSize = table.getRow(i).getTableCells().size();
            for (int j = 0; j < tableCellsSize; j++) {
                if (isOmitCell(i, j)) {
                    continue;
                }
                XWPFTableCell tableCell = table.getRow(i).getCell(j);
                CTTcPr tcPr = tableCell.getCTTc().getTcPr();
                int colspan = getColspan(tcPr);
                if (colspan > 1) { // 合并的列
                    tableToHtmlStr.append("<td colspan='" + colspan + "'");
                } else { // 正常列
                    tableToHtmlStr.append("<td");
                }

                List<Boolean> list = new ArrayList<>();
                getRowspan(table, i, j, list);
                int rowspan = list.size() + 1;
                // System.out.println("第" + i + "行" + "第" + j + "列: " + rowspan);
                if (rowspan > 1) { // 合并的行
                    addOmitCell(i, j, rowspan);
                    tableToHtmlStr.append(" rowspan='" + rowspan + "'>");
                } else {
                    tableToHtmlStr.append(">");
                }
                String text = tableCell.getText();
                tableToHtmlStr.append(text + "</td>");

            }
            tableToHtmlStr.append("</tr>");
        }
        tableToHtmlStr.append("</table>");

        clearTableInfo();

        return tableToHtmlStr.toString();
    }

    public void clearTableInfo() {
        omitCellsList = new ArrayList<>();
        tableRows = 0;
    }

    public static void main(String[] args) {
        ReadWordTable readWordTable = new ReadWordTable();

        try (FileInputStream fileInputStream = new FileInputStream("普通表格.docx");
                XWPFDocument document = new XWPFDocument(fileInputStream);) {
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables) {
                System.out.println(readWordTable.readTable(table));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
