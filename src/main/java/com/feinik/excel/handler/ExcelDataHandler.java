package com.feinik.excel.handler;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel数据处理
 *
 * @author Feinik
 */
public interface ExcelDataHandler {

    /**
     * Excel head头部字体设置
     * @param cellIndex 列索引
     * @return CellStyle
     */
    CellStyle headFont(int cellIndex);

    /**
     * Excel 除head外的内容字体设置
     * @param cellIndex 列索引
     * @param data 行数据对象
     * @return CellStyle
     */
    CellStyle contentFont(int cellIndex, Object data);

    /**
     * Excel sheet
     * @param sheetIndex sheet索引
     * @param sheet
     */
    void sheet(int sheetIndex, Sheet sheet);

    /**
     * workbook context 初始化回调一次
     * @param workbook
     */
    void workbookContext(Workbook workbook);
}
