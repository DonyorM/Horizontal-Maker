package horizontal_maker;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

public class Vertical {

    public static final int TABLE_HEIGHT = 11000;
    public static final BigInteger COL_WIDTH = BigInteger.valueOf(3000);

    //Borrowed from: https://stackoverflow.com/questions/34647624/how-to-colspan-a-table-in-word-with-apache-poi/34663420
    private static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }
    
    private static void mergeCellHorizontall(XWPFTable table, int row, int fromCol, int toCol) {
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            cell.getCTTc().addNewTcPr().addNewHMerge().setVal(colIndex == fromCol ? STMerge.RESTART : STMerge.CONTINUE);
        }
    }

    public static XWPFRun addCellStyle(XWPFTableCell cell, String text) {
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        run.setText(text);
        run.setFontSize(14);
        return run;
    }

    public static void createVerticalSegment(List<String> paragraphTitles,
            String segmentName, int segmentCount, String ref, XWPFDocument doc) {
        int paraCount = paragraphTitles.size();

        //One extra row for the top and one for the application and one for summary
        XWPFTable tbl = doc.createTable(paraCount + 3, 3);
        for (int i = 0; i < 3; i++) {
            // widthCellsAcrossRow(tbl, 0, i, COL_WIDTH);
            tbl.getCTTbl().addNewTblGrid().addNewGridCol().setW(COL_WIDTH);
        }
        List<XWPFTableRow> rows = tbl.getRows();
        XWPFTableRow firstRow = rows.get(0);
        firstRow.setHeight(10);
        addCellStyle(firstRow.getCell(0), "Segment " + segmentCount).setColor("007171");
        addCellStyle(firstRow.getCell(1), segmentName);
        addCellStyle(firstRow.getCell(2), ref);
        firstRow.getCell(0).getParagraphArray(0).setAlignment(ParagraphAlignment.LEFT);
        firstRow.getCell(1).getParagraphArray(0).setAlignment(ParagraphAlignment.CENTER);
        firstRow.getCell(2).getParagraphArray(0).setAlignment(ParagraphAlignment.RIGHT);
        int rowHeight = TABLE_HEIGHT / (rows.size() - 1);
        for (int i = 1; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
            row.setHeight(rowHeight);
            if (i - 1 < paragraphTitles.size()) {
                XWPFRun paraRun = row.getCell(1).getParagraphs().get(0).createRun(); //Second column is where the paragraph titles go
                paraRun.setText(paragraphTitles.get(i - 1));
                paraRun.setUnderline(UnderlinePatterns.SINGLE);
            }
        }
        
        XWPFRun sumRun = tbl.getRow(paraCount + 1).getCell(1).getParagraphArray(0).createRun();
        sumRun.setText("Summary:");
        sumRun.setBold(true);
        
        XWPFRun appRun = tbl.getRow(paraCount + 2).getCell(0).getParagraphArray(0).createRun();
        appRun.setText("Application:");
        appRun.setColor("007826");
        
        //Don't want to merge the header row so start at 1
        mergeCellVertically(tbl, 0, 1, paraCount + 1);
        mergeCellVertically(tbl, 2, 1, paraCount + 1);
        //Merge application row (last row)
        mergeCellHorizontall(tbl, paraCount + 2, 0, 2);
    }

    public static void createDocument(String outputLoc, List<List<String>> paraTitles,
            List<String> names, List<String> refs) {
        XWPFDocument doc = new XWPFDocument();
        for (int i = 0; i < names.size(); i++) {
            createVerticalSegment(paraTitles.get(i), names.get(i), i + 1, refs.get(i), doc);
            doc.createParagraph().setPageBreak(true);
        }
        try (FileOutputStream str = new FileOutputStream(outputLoc)) {
            doc.write(str);
        } catch (IOException ex) {
            Logger.getLogger(Vertical.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
