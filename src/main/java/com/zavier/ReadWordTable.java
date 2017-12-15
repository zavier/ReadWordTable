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

/**
 * 读取word中的表格，包括复杂表格（合并的单元格）
 */
public class ReadWordTable {

    /**
     * 保存需要被忽略的单元格
     */
    private List<String> omitCellsList = new ArrayList<>();

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
     * 获取当前单元格的colspan（列合并）的列数
     * 
     * @param tcPr 单元格属性
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
     * 获取当前单元格的rowspan（行合并）的行数-1（因为没有包括其本身）
     * 
     * @param table 表格
     * @param row 行值
     * @param col 列值
     * @param list list中值的数量表示行合并单元格的数量-1
     */
    public void getRowspan(XWPFTable table, int row, int col, List<Boolean> list) {

        XWPFTableCell cell = table.getRow(row).getCell(col);
        // 正常独立单元格
        if (!isContinueRow(cell) && !isRestartRow(cell)) {
            return;
        }
        // 已达到最后一行
        if (row + 1 >= table.getNumberOfRows()) {
            return;
        }
        // 当前单元格的宽度
        int standWidth = getCellWidth(table, row, col);
        // 当前单元格距离左侧边框的距离
        int standLeftWidth = getLeftWidth(table, row, col);

        row = row + 1;
        int colsNum = table.getRow(row).getTableCells().size();
        // 因为列合并单元格可能导致行合并的单元格并不在同一列，所以从头遍历列，通过属性、宽度以及距离左边框间距来判断是否是行合并
        for (int i = 0; i < colsNum; i++) {
            XWPFTableCell testTable = table.getRow(row).getCell(i);
            // 是否为合并单元格的中间行（包括结尾行）
            if (isContinueRow(testTable)) {
                // 是被上一行单元格合并的单元格
                if (getCellWidth(table, row, i) == standWidth
                        && getLeftWidth(table, row, i) == standLeftWidth) {
                    list.add(true);
                    // 被合并的单元格在生成html时需要忽略
                    addOmitCell(row, i);
                    // 如果下一行仍有此列，则继续查找
                    if (col < colsNum) {
                        getRowspan(table, row, col, list);
                    }
                    break;
                }
            }
        }
    }

    /**
     * 判断是否是合并行的起始行单元格
     * 
     * @param tableCell
     * @return
     */
    public boolean isRestartRow(XWPFTableCell tableCell) {
        CTTcPr tcPr = tableCell.getCTTc().getTcPr();
        if (tcPr.getVMerge() == null) {
            return false;
        }
        if (tcPr.getVMerge().getVal() == null) {
            return false;
        }
        if (tcPr.getVMerge().getVal().toString().equalsIgnoreCase("restart")) {
            return true;
        }
        return false;
    }

    /**
     * 判断是否是合并行的中间行单元格（包括结尾的最后一行的单元格）
     * 
     * @param tableCell
     * @return
     */
    public boolean isContinueRow(XWPFTableCell tableCell) {
        CTTcPr tcPr = tableCell.getCTTc().getTcPr();
        if (tcPr.getVMerge() == null) {
            return false;
        }
        if (tcPr.getVMerge().getVal() == null) {
            return true;
        }
        return false;
    }

    public int getLeftWidth(XWPFTable table, int row, int col) {
        int leftWidth = 0;
        for (int i = 0; i < col; i++) {
            leftWidth += getCellWidth(table, row, i);
        }
        return leftWidth;
    }

    public int getCellWidth(XWPFTable table, int row, int col) {
        BigInteger width = table.getRow(row).getCell(col).getCTTc().getTcPr().getTcW().getW();
        return width.intValue();
    }

    /**
     * 添加忽略的单元格(被行合并的单元格，生成HTML时需要忽略)
     * 
     * @param row
     * @param col
     */
    public void addOmitCell(int row, int col) {
        String omitCellStr = generateOmitCellStr(row, col);
        omitCellsList.add(omitCellStr);
    }

    public boolean isOmitCell(int row, int col) {
        String cellStr = generateOmitCellStr(row, col);
        return omitCellsList.contains(cellStr);
    }

    public String readTable(XWPFTable table) throws IOException {
        // 表格行数
        int tableRowsSize = table.getRows().size();
        StringBuilder tableToHtmlStr = new StringBuilder("<table>");

        for (int i = 0; i < tableRowsSize; i++) {
            tableToHtmlStr.append("<tr>");
            int tableCellsSize = table.getRow(i).getTableCells().size();
            for (int j = 0; j < tableCellsSize; j++) {
                if (isOmitCell(i, j)) {
                    continue;
                }
                XWPFTableCell tableCell = table.getRow(i).getCell(j);
                // 获取单元格的属性
                CTTcPr tcPr = tableCell.getCTTc().getTcPr();
                int colspan = getColspan(tcPr);
                if (colspan > 1) { // 合并的列
                    tableToHtmlStr.append("<td colspan='" + colspan + "'");
                } else { // 正常列
                    tableToHtmlStr.append("<td");
                }

                // list 的长度表示当前单元格行合并的单元格数-1
                List<Boolean> list = new ArrayList<>();
                getRowspan(table, i, j, list);
                int rowspan = list.size() + 1;
                // System.out.println("第" + i + "行" + "第" + j + "列: " + rowspan);
                if (rowspan > 1) { // 合并的行
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
        // System.out.println(omitCellsList);
        omitCellsList.clear();
    }

    public static void main(String[] args) {
        ReadWordTable readWordTable = new ReadWordTable();

        try (FileInputStream fileInputStream = new FileInputStream("table1.docx");
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
