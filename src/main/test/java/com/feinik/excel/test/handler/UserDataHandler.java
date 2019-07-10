package com.feinik.excel.test.handler;

import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.test.model.UserData;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author Feinik
 */
public class UserDataHandler implements ExcelDataHandler {

    @Override
    public void headCellStyle(CellStyle style, int cellIndex) {
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
    }

    @Override
    public void headFont(Font font, int cellIndex) {
        font.setColor(IndexedColors.WHITE.getIndex());
    }

    @Override
    public void contentCellStyle(CellStyle style, int cellIndex) {
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    @Override
    public void contentFont(Font font, int cellIndex, Object data) {
        UserData user = (UserData) data;
        switch (cellIndex) {
            case 2: //这里的值为Model对象中ExcelProperty注解里的index值
                if (Integer.valueOf(user.getAge()) > 60) { //表示将年龄大于60的第二列也就是工资列的cell字体标记为红色
                    font.setColor(IndexedColors.RED.getIndex());
                    font.setFontName("宋体");
                    font.setItalic(true);
                    font.setBold(true);
                }
                break;

        }
    }

    @Override
    public void sheet(int sheetIndex, Sheet sheet) {}
}
