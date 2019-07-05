package com.feinik.excel.handler;

import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;

/**
 * Write excel style handler
 *
 * @author Feinik
 */
public class StyleHandler implements WriteHandler {
    private static final Logger log = LoggerFactory.getLogger(StyleHandler.class);
    private Map<String, List<? extends BaseRowModel>> dataMap;
    private List<? extends BaseRowModel> data;
    private ExcelDataHandler handler;
    private int rowIndex;

    public void setHandler(ExcelDataHandler handler) {
        this.handler = handler;
    }

    public StyleHandler(Map<String, List<? extends BaseRowModel>> dataMap) {
        this.dataMap = dataMap;
    }

    @Override
    public void sheet(int sheetIndex, Sheet sheet) {
        data = dataMap.get(sheet.getSheetName());
        handler.sheet(sheetIndex, sheet);
    }

    @Override
    public void row(int rowIndex, Row row) {
        this.rowIndex = rowIndex;
    }

    @Override
    public void cell(int cellIndex, Cell cell) {
        final Workbook workbook = cell.getSheet().getWorkbook();
        final CellStyle style = createStyle(workbook);
        final Font font = workbook.createFont();
        try {
            if (handler != null) {
                if (rowIndex == 0) {
                    handler.headFont(font, cellIndex);
                    handler.headCellStyle(style, cellIndex);

                } else {
                    handler.contentFont(font, cellIndex, data.get(rowIndex-1));
                    handler.contentCellStyle(style, cellIndex);
                }
            }

        } catch (Exception e) {
            log.error("excel cell process failed, cause is:{}", e);
        }
        style.setFont(font);
        cell.getRow().getCell(cellIndex).setCellStyle(style);

    }

    private CellStyle createStyle(Workbook workbook) {
        final CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }
}
