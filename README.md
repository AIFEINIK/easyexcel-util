# easyexcel-util
本项目基于阿里easyexcel，使其更容易处理每个cell的字体与样式  

# 自定义注解 xcelValueFormat  
通过该注解更加方便的处理每个数据的具体格式  
```
public class UserData extends BaseRowModel implements Serializable {

    @ExcelProperty(value = "用户名", index = 0)
    private String userName;

    @ExcelProperty(value = "年龄", index = 1)
    private Integer age;

    @ExcelProperty(value = "工资", index = 2)
    @ExcelValueFormat(format = "{0}￥")
    private String salary;

}
```

# 通过实现 xcelDataHandler 接口来设置具体每个cell的样式与字体，如：
```
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
            case 2: //这里的值与Model对象中 @ExcelProperty(value = "用户名", index = 0)注解里的index值
                if (Integer.valueOf(user.getAge()) > 60) {
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
```

