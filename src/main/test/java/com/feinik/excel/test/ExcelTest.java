package com.feinik.excel.test;

import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.feinik.excel.test.listener.ExcelListener;
import com.feinik.excel.test.util.FileUtil;
import com.feinik.excel.utils.ExcelUtil;
import org.junit.Test;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author Feinik
 */
public class ExcelTest {

    @Test
    public void writeTest() {
        UserData userData = new UserData();
        userData.setUserName("张三");
        userData.setAge(62);
        userData.setSalary("5000");

        UserData userData2 = new UserData();
        userData2.setUserName("李四");
        userData2.setAge(55);
        userData2.setSalary("7000");

        List<UserData> data = new ArrayList<>();
        data.add(userData);

        List<UserData> data2 = new ArrayList<>();
        data2.add(userData2);

        Map<String, List<? extends BaseRowModel>> map = new HashMap<>();
        map.put("sheet1", data);
        map.put("sheet2", data2);

        try {
            //将数据写入单个sheet
            ExcelUtil.writeExcelWithOneSheet(new File("G:/tmp/test.xlsx"),
                    "用户信息",
                    data);

            //将数据写入单个sheet, 并通过实现ExcelDataHandler接口来指定具体excell的样式
            //ExcelUtil.writeExcelWithOneSheet(new File("G:/tmp/test.xlsx"),
            //        "用户信息",
            //        data,
            //        new com.feinik.excel.test.UserDataHandler());

            //将数据写入多个sheet
            //ExcelUtil.writeExcelWithMultiSheet(new File("G:/tmp/test.xlsx"),
            //        map);

            //将数据写入多个sheet, 并通过实现ExcelDataHandler接口来指定具体excell的样式
            //ExcelUtil.writeExcelWithMultiSheet(new File("G:/tmp/test.xlsx"),
            //        map,
            //        new UserDataHandler());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readSmallFilesTest() {
        try (InputStream in = FileUtil.getResourcesFileInputStream("test1.xlsx")) {
            final List<Object> data = ExcelUtil.read(in, new Sheet(1, 1));
            print(data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readSmallFilesCastModelTest() {
        try (InputStream in = FileUtil.getResourcesFileInputStream("test1.xlsx")) {
            final List<Object> data = ExcelUtil.read(in, new Sheet(1, 1, UserData.class));
            print(data);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void readLargeFilesTest() {
        try (InputStream in = FileUtil.getResourcesFileInputStream("test1.xlsx")) {
            ExcelListener listener = new ExcelListener();
            ExcelUtil.readBySax(in, new Sheet(1, 1), listener);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void print(List<Object> datas) {
        int i = 0;
        for (Object ob : datas) {
            System.out.println(i++);
            System.out.println(ob);
        }
    }
}
