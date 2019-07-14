package com.feinik.excel.test;

import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.feinik.excel.ExcelWrapWriter;
import com.feinik.excel.test.handler.CampaignDataHandler;
import com.feinik.excel.test.listener.ExcelListener;
import com.feinik.excel.test.model.CampaignModel;
import com.feinik.excel.test.util.FileUtil;
import com.feinik.excel.utils.ExcelUtil;
import com.google.common.collect.Lists;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Feinik
 */
public class ExcelTest {

    CampaignModel m1 = new CampaignModel("2019-01-01", "10000000", "campaign1", "12.21", "100", "0.11");
    CampaignModel m2 = new CampaignModel("2019-01-02", "12000010", "campaign2", "13", "99", "0.91");
    CampaignModel m3 = new CampaignModel("2019-01-03", "12001010", "campaign3", "10", "210", "1.13");
    CampaignModel m4 = new CampaignModel("2019-01-04", "15005010", "campaign4", "21.9", "150", "0.15");

    ArrayList<CampaignModel> data1 = Lists.newArrayList(m1, m2);
    ArrayList<CampaignModel> data2 = Lists.newArrayList(m3, m4);

    /**
     * 小数据量一次性写入单个sheet，使用默认样式
     *
     * @throws Exception
     */
    @Test
    public void writeExcelWithOneSheet() throws Exception {
        ExcelUtil.writeExcelWithOneSheet(new File("G:/tmp/campaign.xlsx"),
                "campaign",true, data1);
    }

    /**
     * 小数据量一次性写入单个sheet，使用自定义样式
     *
     * @throws Exception
     */
    @Test
    public void writeExcelWithOneSheet2() throws Exception {
        ExcelUtil.writeExcelWithOneSheet(new File("G:/tmp/campaign.xlsx"),
                "campaign",true, data1, new CampaignDataHandler());
    }

    /**
     * 小数据量一次性写入多个sheet，默认样式
     *
     * @throws Exception
     */
    @Test
    public void writeExcelWithMultiSheet() throws Exception {
        Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
        map.put("sheet1", data1);
        map.put("sheet2", data2);

        ExcelUtil.writeExcelWithMultiSheet(new File("G:/tmp/campaign.xlsx"), map, true);
    }

    /**
     * 小数据量一次性写入多个sheet，使用自定义样式
     *
     * @throws Exception
     */
    @Test
    public void writeExcelWithMultiSheet2() throws Exception {
        Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
        map.put("sheet1", data1);
        map.put("sheet2", data2);

        ExcelUtil.writeExcelWithMultiSheet(new File("G:/tmp/campaign.xlsx"),
                map,true, new CampaignDataHandler());
    }

    /**
     * 单个sheet
     * 测试分批写入excel文件，可通过该方式写入超大数据，而不至于一次写入大数据量导致OOM问题
     */
    @Test
    public void writeOneSheetWithWrapWriter() {
        ExcelWrapWriter wrapWriter = null;
        try {
            OutputStream os = new FileOutputStream("G:/tmp/campaign.xlsx");
            //默认样式
            //wrapWriter = new ExcelWrapWriter(os, ExcelTypeEnum.XLSX);

            //自定义excel样式
            wrapWriter = new ExcelWrapWriter(os, ExcelTypeEnum.XLSX, new CampaignDataHandler());

            List<CampaignModel> models1 = Lists.newArrayList(m1, m2);
            List<CampaignModel> models2 = Lists.newArrayList(m3, m4);

            //第一批次写入设置包含head头
            ExcelUtil.writeExcelWithOneSheet(wrapWriter, "sheet1", true, models1);

            //第二批次开始不需要在写入head头
            ExcelUtil.writeExcelWithOneSheet(wrapWriter, "sheet1", false, models2);
        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            //close IO
            if (wrapWriter != null) {
                wrapWriter.finish();
            }
        }
    }

    /**
     * 多个sheet
     * 测试分批写入excel文件，可通过该方式写入超大数据，而不至于一次写入大数据量导致OOM问题
     */
    @Test
    public void writeMultiSheetWithWrapWriter() {
        ExcelWrapWriter wrapWriter = null;
        try {
            //os流不需要单独close，可通过wrapWriter.finish()来关闭
            OutputStream os = new FileOutputStream("G:/tmp/campaign.xlsx");
            //默认样式
            //wrapWriter = new ExcelWrapWriter(os, ExcelTypeEnum.XLSX);

            //自定义excel样式
            wrapWriter = new ExcelWrapWriter(os, ExcelTypeEnum.XLSX, new CampaignDataHandler());
            Map<String, List<? extends BaseRowModel>> batch1 = new HashMap<>();
            List<CampaignModel> models1 = Lists.newArrayList(m1, m2);
            List<CampaignModel> models2 = Lists.newArrayList(m3, m4);
            batch1.put("sheet1", models1);
            batch1.put("sheet2", models2);

            Map<String, List<? extends BaseRowModel>> batch2 = new HashMap<>();
            List<CampaignModel> models3 = Lists.newArrayList(m4, m2);
            List<CampaignModel> models4 = Lists.newArrayList(m3, m1);
            batch2.put("sheet1", models3);
            batch2.put("sheet2", models4);

            //第一批次写入设置包含head头
            ExcelUtil.writeExcelWithMultiSheet(wrapWriter, true, batch1);

            //第二批次开始不需要在写入head头
            ExcelUtil.writeExcelWithMultiSheet(wrapWriter, false, batch2);
        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            //close IO
            if (wrapWriter != null) {
                wrapWriter.finish();
            }
        }
    }

    @Test
    public void readSmallFilesTest() {
        try (InputStream in = FileUtil.getResourcesFileInputStream("campaign.xlsx")) {
            final List<Object> data = ExcelUtil.read(in, new Sheet(1, 1));
            print(data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readSmallFilesCastModelTest() {
        try (InputStream in = FileUtil.getResourcesFileInputStream("campaign.xlsx")) {
            final List<Object> data = ExcelUtil.read(in, new Sheet(1, 1, CampaignModel.class));
            print(data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readLargeFilesTest() {
        try (InputStream in = FileUtil.getResourcesFileInputStream("campaign.xlsx")) {
            ExcelListener listener = new ExcelListener();
            ExcelUtil.readBySax(in, new Sheet(1, 1), listener);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void print(List<Object> datas) {
        for (Object ob : datas) {
            System.out.println(ob);
        }
    }
}
