# easyexcel-util
本项目基于阿里easyexcel，使其更容易处理每个cell的字体与样式  

# Maven包引入
```
<dependency>
    <groupId>com.github.aifeinik</groupId>
    <artifactId>easyexcel-util</artifactId>
    <version>0.1.1</version>
</dependency>

```

# 自定义注解 ExcelValueFormat  
通过该注解更加方便的处理每个数据的具体格式, 内部采用MessageFormat.format进行数据格式化
```
@Data
public class CampaignModel extends BaseRowModel implements Serializable {

    @ExcelProperty(value = "日期", index = 0)
    private String day;

    @ExcelProperty(value = "广告系列 ID", index = 1)
    private String campaignId;

    @ExcelProperty(value = "广告系列", index = 2)
    private String campaignName;

    @ExcelProperty(value = "费用", index = 3)
    @ExcelValueFormat(format = "{0}$")
    private String cost;

    @ExcelProperty(value = "点击次数", index = 4)
    private String clicks;

    @ExcelProperty(value = "点击率", index = 5)
    @ExcelValueFormat(format = "{0}%")
    private String ctr;

}
```

# 通过实现 ExcelDataHandler 接口来设置具体每个cell的样式与字体，如：
```
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

public class CampaignDataHandler implements ExcelDataHandler {

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
        CampaignModel campaign = (CampaignModel) data;
        switch (cellIndex) {
            case 4: //这里的值为Model对象中ExcelProperty注解里的index值
                if (Long.valueOf(campaign.getClicks()) > 100) { //表示将点击次数大于100的第4列也就是点击次数列的cell字体标记为红色
                    font.setColor(IndexedColors.RED.getIndex());
                    font.setFontName("宋体");
                    font.setItalic(true);
                    font.setBold(true);
                }
                break;

        }
    }

    @Override
    public void sheet(int sheetIndex, Sheet sheet) {
        System.out.println("sheetIndex = [" + sheetIndex + "]");
    }
}
```
# Excel数据写入
## 小数据量一次性写入单个sheet，使用默认样式
```
public class ExcelTest {

    CampaignModel m1 = new CampaignModel("2019-01-01", "10000000", "campaign1", "12.21", "100", "0.11");
    CampaignModel m2 = new CampaignModel("2019-01-02", "12000010", "campaign2", "13", "99", "0.91");
    CampaignModel m3 = new CampaignModel("2019-01-03", "12001010", "campaign3", "10", "210", "1.13");
    CampaignModel m4 = new CampaignModel("2019-01-04", "15005010", "campaign4", "21.9", "150", "0.15");

    ArrayList<CampaignModel> data1 = Lists.newArrayList(m1, m2);
    ArrayList<CampaignModel> data2 = Lists.newArrayList(m3, m4);

    @Test
    public void writeExcelWithOneSheet() throws Exception {
        ExcelUtil.writeExcelWithOneSheet(new File("G:/tmp/campaign.xlsx"),
                "campaign",
                data1);
    }
}
```
![s1](https://github.com/AIFEINIK/img-resource/blob/master/easyexcel-util/0011.png)

## 小数据量一次性写入单个sheet，使用自定义样式
```
    @Test
    public void writeExcelWithOneSheet2() throws Exception {
        ExcelUtil.writeExcelWithOneSheet(new File("G:/tmp/campaign.xlsx"),
                "campaign",
                data1,
                new CampaignDataHandler());
    }
```
![s2](https://github.com/AIFEINIK/img-resource/blob/master/easyexcel-util/0012.png)

## 小数据量一次性写入多个sheet，默认样式
```
    @Test
    public void writeExcelWithMultiSheet() throws Exception {
        Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
        map.put("sheet1", data1);
        map.put("sheet2", data2);

        ExcelUtil.writeExcelWithMultiSheet(new File("G:/tmp/campaign.xlsx"), map);
    }
```
![s3](https://github.com/AIFEINIK/img-resource/blob/master/easyexcel-util/0013.png)

## 小数据量一次性写入多个sheet，使用自定义样式
```
    @Test
    public void writeExcelWithMultiSheet2() throws Exception {
        Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
        map.put("sheet1", data1);
        map.put("sheet2", data2);

        ExcelUtil.writeExcelWithMultiSheet(new File("G:/tmp/campaign.xlsx"), map, new CampaignDataHandler());
    }
```
![s3](https://github.com/AIFEINIK/img-resource/blob/master/easyexcel-util/0014.png)

# 测试代码
[ExcelTest](https://github.com/AIFEINIK/easyexcel-util/blob/master/src/main/test/java/com/feinik/excel/test/ExcelTest.java)
