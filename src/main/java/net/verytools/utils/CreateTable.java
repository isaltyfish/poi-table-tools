package net.verytools.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

public class CreateTable {

    private static final int TWIPS_PER_INCH = 1440;

    public static void main(String[] args) throws IOException {
        try (XWPFDocument doc = new XWPFDocument()) {
            try (FileOutputStream fos = new FileOutputStream("D:\\tmp\\table.docx")) {
                XWPFTable table = doc.createTable(1, 2);

                // grid for table
                CTTblGrid grid = table.getCTTbl().addNewTblGrid();
                grid.addNewGridCol().setW(BigInteger.valueOf((long) (4 / 2.54 * TWIPS_PER_INCH)));
                grid.addNewGridCol().setW(BigInteger.valueOf((long) (8 / 2.54 * TWIPS_PER_INCH)));

                table.setInsideHBorder(XWPFTable.XWPFBorderType.DOTTED, 4, 0, "000000");
                table.setInsideVBorder(XWPFTable.XWPFBorderType.DOTTED, 4, 0, "000000");
                table.setLeftBorder(XWPFTable.XWPFBorderType.DOTTED, 4, 0, "000000");
                table.setRightBorder(XWPFTable.XWPFBorderType.DOTTED, 4, 0, "000000");
                table.setTopBorder(XWPFTable.XWPFBorderType.DOTTED, 4, 0, "000000");
                table.setBottomBorder(XWPFTable.XWPFBorderType.DOTTED, 4, 0, "000000");

                // should use getRow instead of createRow
                XWPFTableRow row1 = table.getRow(0);

                // set row height to 2cm, but we have to convert it to twips
                row1.setHeight((int) (2 / 2.54 * 1440));
//                row1.setHeightRule(TableRowHeightRule.AT_LEAST);

                // set text to cell
                row1.getCell(0).setText("row 1 cell 1");
                row1.getCell(1).setText("row 1 cell 2");
                row1.getCell(0).setWidth(Long.toString((long) (4 / 2.54 * TWIPS_PER_INCH)));
                row1.getCell(1).setWidth(Long.toString((long) (8 / 2.54 * TWIPS_PER_INCH)));

                // fill cells with orange color
                row1.getCell(0).setColor("ffa500");
                row1.getCell(1).setColor("ffa500");

                row1.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                row1.getCell(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);

                XWPFTableRow row2 = table.createRow();
                row2.getCell(0).setText("row 2 cell 1");
                row2.getCell(1).setText("row 2 cell 2");

                // 合并单元格
                TableTools.mergeCellsHorizontal(table, 0, 0, 1);
//                CellMergeUtil.mergeCellHorizontally(table, 0, 0, 1);

                doc.write(fos);
            }
        }
    }

}
