package com.feinik.excel.handler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Excel数据处理
 *
 * @author Feinik
 */
public interface ExcelDataHandler {

    /**
     * Excel head头部字体设置
     * @param font
     * @param cellIndex 列索引
     */
    void headFont(Font font, int cellIndex);

    /**
     * Excel head头部样式设置
     * @param style
     * @param cellIndex 列索引
     */
    void headCellStyle(CellStyle style, int cellIndex);

    /**
     * Excel 除head外的内容字体设置
     * @param font
     * @param cellIndex 列索引
     */
    void contentFont(Font font, int cellIndex, Object data);

    /**
     * Excel 除head外的内容样式设置
     * @param style
     * @param cellIndex 列索引
     */
    void contentCellStyle(CellStyle style, int cellIndex);

    /**
     * Excel sheet
     * @param sheetIndex sheet索引
     * @param sheet
     */
    void sheet(int sheetIndex, Sheet sheet);
}
