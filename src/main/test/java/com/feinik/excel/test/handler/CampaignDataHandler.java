package com.feinik.excel.test.handler;

import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.test.model.CampaignModel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 *
 * @author Feinik
 */
public class CampaignDataHandler implements ExcelDataHandler {

    //通过对象池的方式来解决大数据量下重复创建CellStyle而导致的异常
    //相同设置属性的CellStyle对象达到可复用
    private Map<String, CellStyle> stylePool = new HashMap<>();

    private enum ObjType {
        HEAD_STYLE,
        INDEX4_STYLE,
        DEFAULT_STYLE
    }

    @Override
    public void workbookContext(Workbook workbook) {
        putStylePool(workbook);
    }

    //初始化对象池
    private void putStylePool(Workbook workbook) {
        stylePool.put(ObjType.HEAD_STYLE.name(), setAndGetHeadStyle(workbook));
        stylePool.put(ObjType.INDEX4_STYLE.name(), setAndGetIndex4Style(workbook));
        stylePool.put(ObjType.DEFAULT_STYLE.name(), setAndGetDefaultStyle(workbook));
    }

    @Override
    public CellStyle headFont(int cellIndex) {
        //从对象池直接获取
        return stylePool.get(ObjType.HEAD_STYLE.name());
    }

    @Override
    public CellStyle contentFont(int cellIndex, Object data) {
        CampaignModel campaign = (CampaignModel) data;

        switch (cellIndex) {
            case 4: //这里的值为Model对象中ExcelProperty注解里的index值
                if (Long.valueOf(campaign.getClicks()) > 100) { //表示将点击次数大于100的第4列也就是点击次数列的cell字体标记为红色
                    return stylePool.get(ObjType.INDEX4_STYLE.name());

                } else {
                    return stylePool.get(ObjType.DEFAULT_STYLE.name());
                }

            default:
                return stylePool.get(ObjType.DEFAULT_STYLE.name());

        }
    }

    @Override
    public void sheet(int sheetIndex, Sheet sheet) {
        System.out.println("sheetIndex = [" + sheetIndex + "]");
    }

    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    public CellStyle setAndGetHeadStyle(Workbook book) {
        final Font font = book.createFont();
        final CellStyle style = getCellStyle(book);

        font.setColor(IndexedColors.WHITE.getIndex());
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        font.setFontName("微软雅黑");

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        style.setFont(font);
        return style;
    }

    public CellStyle setAndGetDefaultStyle(Workbook book) {
        final Font font = book.createFont();
        final CellStyle style = getCellStyle(book);
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        return style;
    }

    public CellStyle setAndGetIndex4Style(Workbook book) {
        final Font font = book.createFont();
        final CellStyle style = getCellStyle(book);
        font.setColor(IndexedColors.RED.getIndex());
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        return style;
    }
}
