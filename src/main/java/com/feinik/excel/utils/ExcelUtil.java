package com.feinik.excel.utils;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.feinik.excel.annotation.ExcelValueFormat;
import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.handler.StyleHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * ExcelUtil
 *
 * @author Feinik
 */
@Slf4j
public class ExcelUtil {

    /**
     * Write one sheet
     * @param writeFile Writes data to the specified file
     * @param sheetName sheet name
     * @param data Write data
     * @param handler The data processor, through which you can specify a specific style for each cell
     * @return write result
     */
    public static boolean writeExcelWithOneSheet(File writeFile,
                                              String sheetName,
                                              List<? extends BaseRowModel> data,
                                              ExcelDataHandler handler) throws Exception {
        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        StyleHandler sh = new StyleHandler(dataMap);
        sh.setHandler(handler);
        return writeExcel(writeFile,  sheetName, data, sh);
    }

    /**
     * Write one sheet
     * @param writeFile Writes data to the specified file
     * @param sheetName sheet name
     * @param data Write data
     * @return write result
     */
    public static boolean writeExcelWithOneSheet(File writeFile,
                                              String sheetName,
                                              List<? extends BaseRowModel> data) throws Exception {
        return writeExcel(writeFile,  sheetName, data, null);
    }

    /**
     * Write multiple sheets
     * @param writeFile Writes data to the specified file
     * @param data Write data
     * @param handler The data processor, through which you can specify a specific style for each cell
     * @return write result
     */
    public static boolean writeExcelWithMultiSheet(File writeFile,
                                                Map<String, List<? extends BaseRowModel>> data,
                                                ExcelDataHandler handler) throws Exception {
        StyleHandler sh = new StyleHandler(data);
        sh.setHandler(handler);
        return writeExcelWithHandler(writeFile, data, sh);
    }

    /**
     * Write multiple sheets
     * @param writeFile Writes data to the specified file
     * @param data Write data
     * @return write result
     */
    public static boolean writeExcelWithMultiSheet(File writeFile,
                                                Map<String, List<? extends BaseRowModel>> data) throws Exception {
        return writeExcelWithHandler(writeFile, data, null);
    }

    private static boolean writeExcel(File writeFile,
                                   String sheetName,
                                   List<? extends BaseRowModel> data,
                                   WriteHandler handler) throws Exception {
        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        return writeExcelWithHandler(writeFile,  dataMap, handler);
    }

    /**
     * Write excel
     * @param writeFile Writes data to the specified file
     * @param data Write data
     * @param handler
     * @return write result
     */
    private static boolean writeExcelWithHandler(File writeFile,
                                              Map<String, List<? extends BaseRowModel>> data,
                                              WriteHandler handler) throws Exception {
        boolean result = false;
        if (MapUtils.isEmpty(data)) {
            log.warn("write excel data not empty");
            return false;
        }

        try (OutputStream os = new FileOutputStream(writeFile)) {
            ExcelWriter ew = EasyExcelFactory.getWriterWithTempAndHandler(null, os, ExcelTypeEnum.XLSX, true, handler);

            int sheetIndex = 1;
            for (String sheetName : data.keySet()) {
                List<? extends BaseRowModel> models = data.get(sheetName);
                if (models.size() == 0) {
                    continue;
                }
                models = convertData(models);

                Sheet sheet = new Sheet(sheetIndex, 0, models.get(0).getClass());
                sheetName = StringUtils.isEmpty(sheetName) ? "sheet" + sheetIndex : sheetName;
                sheet.setSheetName(sheetName);
                ew.write(models, sheet);
                sheetIndex++;
            }
            ew.finish();
            result = true;
            log.info("write excel data success, file path with:{}", writeFile.getPath());
        } catch (Exception e) {
            log.error("write excel data failed , file path with:{}, cause is:{}", writeFile.getPath(), e);
        }
        return result;
    }

    /**
     * Data formatting
     * @param data
     * @throws Exception
     */
    private static <T> List<T> convertData(List<? extends BaseRowModel> data) throws Exception {
        List<T> result = new ArrayList<>();
        for (Object o : data) {
            final Object copyObj = BeanUtils.transform(o.getClass(), o);

            List<Field> fields = new ArrayList<>() ;
            Class<?> copyObjClass = copyObj.getClass();
            while (copyObjClass != null) {
                fields.addAll(Arrays.asList(copyObjClass.getDeclaredFields()));
                copyObjClass = copyObjClass.getSuperclass();
            }

            for (Field field : fields) {
                field.setAccessible(true);
                final ExcelValueFormat valueFormat = field.getDeclaredAnnotation(ExcelValueFormat.class);
                if (valueFormat != null) {
                    final String format = valueFormat.format();
                    final Object value = field.get(copyObj);
                    if (value == null) {
                        field.set(copyObj, StringUtils.EMPTY);
                    } else {
                        final String newValue = MessageFormat.format(format, value);
                        field.set(copyObj, newValue);
                    }
                }
            }
            result.add((T) copyObj);
        }
        return result;
    }

    /**
     * Quickly read small files，no more than 10,000 lines.
     * @param in
     * @param sheet
     * @return read data
     */
    public static List<Object> read(InputStream in, Sheet sheet) {
        return EasyExcelFactory.read(in, sheet);
    }

    /**
     * Parsing large file
     * @param in
     * @param sheet
     * @param listener Callback method after each row is parsed
     */
    public static void readBySax(InputStream in, Sheet sheet, AnalysisEventListener listener) {
        EasyExcelFactory.readBySax(in, sheet, listener);
    }

    /**
     * Get ExcelReader, All the sheet data can be obtained through ExcelReader object
     * @param in
     * @param listener Callback method after each row is parsed
     * @return ExcelReader
     */
    public static ExcelReader getReader(InputStream in, AnalysisEventListener listener) {
        return EasyExcelFactory.getReader(in, listener);
    }
}
