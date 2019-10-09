package com.deepoove.poi.util;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.awt.*;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class WordTableUtil {

    public static void copyRun(XWPFRun target, XWPFRun source) {
        // 设置run属性
        target.getCTR().setRPr(source.getCTR().getRPr());
        // 设置文本
        target.setText(source.text());
        // 处理图片
        List<XWPFPicture> pictures = source.getEmbeddedPictures();

        for (XWPFPicture picture : pictures) {
            try {
                copyPicture(target, picture);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void copyPicture(XWPFRun target, XWPFPicture picture)throws IOException, InvalidFormatException {

        String filename = picture.getPictureData().getFileName();
        InputStream pictureData = new ByteArrayInputStream(picture
                .getPictureData().getData());
        int pictureType = picture.getPictureData().getPictureType();
        int width = (int) picture.getCTPicture().getSpPr().getXfrm().getExt()
                .getCx();

        int height = (int) picture.getCTPicture().getSpPr().getXfrm().getExt()
                .getCy();

        // target.addBreak();
        target.addPicture(pictureData, pictureType, filename, width, height);
        // target.addBreak(BreakType.PAGE);
    }

    /**
     * 复制段落，从source到target
     * @param target
     * @param source
     *
     */
    public static void copyParagraph(XWPFParagraph target, XWPFParagraph source) {

        // 设置段落样式
        target.getCTP().setPPr(source.getCTP().getPPr());

        // 移除所有的run
        for (int pos = target.getRuns().size() - 1; pos >= 0; pos--) {
            target.removeRun(pos);
        }

        // copy 新的run
        for (XWPFRun s : source.getRuns()) {
            XWPFRun targetrun = target.createRun();
            copyRun(targetrun, s);
        }

    }

    /**
     * 复制单元格，从source到target
     * @param target
     * @param source
     *
     */
    public static void copyTableCell(XWPFTableCell target, XWPFTableCell source) {
        // 列属性
        if (source.getCTTc() != null) {
            target.getCTTc().setTcPr(source.getCTTc().getTcPr());
        }
        // 删除段落
        for (int pos = 0; pos < target.getParagraphs().size(); pos++) {
            target.removeParagraph(pos);
        }
        // 添加段落
        for (XWPFParagraph sp : source.getParagraphs()) {
            XWPFParagraph targetP = target.addParagraph();
            copyParagraph(targetP, sp);
        }
    }

    /**
     *
     * 复制行，从source到target
     * @param target
     * @param source
     *
     */
    public static void copyTableRow(XWPFTableRow target, XWPFTableRow source) {
        // 复制样式
        if (source.getCtRow() != null) {
            target.getCtRow().setTrPr(source.getCtRow().getTrPr());
        }
        // 复制单元格
        for (int i = 0; i < source.getTableCells().size(); i++) {
            XWPFTableCell cell1 = target.getCell(i);
            XWPFTableCell cell2 = source.getCell(i);
            if (cell1 == null) {
                cell1 = target.addNewTableCell();
            }
            copyTableCell(cell1, cell2);
        }
    }
    /**
     * 复制表，从source到target
     * @param target
     * @param source
     */
    public static void copyTable(XWPFTable target, XWPFTable source) {
        // 表格属性
        target.getCTTbl().setTblPr(source.getCTTbl().getTblPr());

        // 复制行
        for (int i = 0; i < source.getRows().size(); i++) {
            XWPFTableRow row1 = target.getRow(i);
            XWPFTableRow row2 = source.getRow(i);
            if (row1 == null) {
                target.addRow(row2);
            } else {
                copyTableRow(row1, row2);
            }
        }
    }
    //此方法必须等全部设置完毕后，才能table。addrow，不然不显示
    public static XWPFTableRow createNewRow(XWPFTable table, XWPFTableRow toClone) throws IOException, XmlException {
        CTRow ctrow = CTRow.Factory.parse(toClone.getCtRow().newInputStream());
        if (table == null) {
            table = toClone.getTable();
        }
        return new XWPFTableRow(ctrow, table);
    }

    public static void setContent(XWPFTableCell cell, Color color, String content){
        XWPFParagraph xwpfParagraph = cell.getParagraphs().get(0);
        XWPFRun xwpfRun = xwpfParagraph.createRun();
        if (color != null) {
            xwpfRun.setColor(ColorUtil.Color2String(color));
        }
        content = content.trim();
        int pos = 0;
        for (char c : content.toCharArray()) {
            if (c == '\n') {
                xwpfRun.addBreak();
                continue;
            } else if (c == '\r') {
                continue;
            }

            xwpfRun.setText(c + "", pos++);
        }
    }
}
