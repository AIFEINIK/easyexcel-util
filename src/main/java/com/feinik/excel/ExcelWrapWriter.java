package com.feinik.excel;

import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.handler.StyleHandler;

import java.io.OutputStream;
import java.util.Map;

/**
 * @author Feinik
 * @discription
 * @date 2019/7/12
 * @since 1.0.0
 */
public class ExcelWrapWriter {

    private ExcelWriter writer;

    private Map<String, Sheet> sheetMap;

    private StyleHandler styleHandler;

    public ExcelWrapWriter(OutputStream os, ExcelTypeEnum excelType) {
        this.writer = EasyExcelFactory.getWriterWithTempAndHandler(null, os, excelType, false, null);
    }

    public ExcelWrapWriter(OutputStream os, ExcelTypeEnum excelType, ExcelDataHandler handler) {
        if (handler == null) {
            throw new IllegalArgumentException("Parameter handler can not be null");
        }
        StyleHandler sh = new StyleHandler();
        sh.setHandler(handler);
        this.styleHandler = sh;
        this.writer = EasyExcelFactory.getWriterWithTempAndHandler(null, os, excelType, false, sh);
    }

    public StyleHandler getStyleHandler() {
        return styleHandler;
    }

    public ExcelWriter getWriter() {
        return writer;
    }

    public Map<String, Sheet> getSheetMap() {
        return sheetMap;
    }

    public void setSheetMap(Map<String, Sheet> sheetMap) {
        this.sheetMap = sheetMap;
    }

    /**
     * Close IO
     */
    public void finish() {
        writer.finish();
    }
}
