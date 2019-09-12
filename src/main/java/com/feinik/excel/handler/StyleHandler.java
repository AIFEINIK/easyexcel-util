package com.feinik.excel.handler;

import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
    private int currentRowIndex;
    private int dataIndex;
    private Workbook workbook;

    public void setHandler(ExcelDataHandler handler) {
        this.handler = handler;
    }

    public StyleHandler(Map<String, List<? extends BaseRowModel>> dataMap) {
        this.dataMap = dataMap;
    }

    public StyleHandler() {
    }

    public void setDataMap(Map<String, List<? extends BaseRowModel>> dataMap) {
        this.dataMap = dataMap;
    }

    @Override
    public void sheet(int sheetIndex, Sheet sheet) {
        this.dataIndex = 0;
        this.data = dataMap.get(sheet.getSheetName());
        this.handler.sheet(sheetIndex, sheet);
    }

    @Override
    public void row(int rowIndex, Row row) {
        if (workbook == null) {
            this.workbook = row.getSheet().getWorkbook();
            handler.workbookContext(workbook);
        }
        this.currentRowIndex = rowIndex;
        this.dataIndex++;
    }

    @Override
    public void cell(int cellIndex, Cell cell) {
        try {
            CellStyle style = null;
            if (handler != null) {
                if (currentRowIndex == 0) {
                    style = handler.headFont(cellIndex);

                } else {
                    style = handler.contentFont(cellIndex, data.get(dataIndex - 1));

                }
            }
            cell.getRow().getCell(cellIndex).setCellStyle(style);

        } catch (Exception e) {
            log.error("excel cell process failed, cause is:{}", e);
        }

    }
}
