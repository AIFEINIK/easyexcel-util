package com.feinik.excel.utils;

import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.feinik.excel.EasyExcelFactory;
import com.feinik.excel.ExcelWrapWriter;
import com.feinik.excel.ExcelWriter;
import com.feinik.excel.annotation.ExcelValueFormat;
import com.feinik.excel.handler.ExcelDataHandler;
import com.feinik.excel.handler.StyleHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
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
     * Write one sheet, Suitable for writing small data volumes
     *
     * @param writeFile Writes data to the specified file
     * @param sheetName sheet name
     * @param needHead  Whether to write header
     * @param data      Write data
     * @param handler   The data processor, through which you can specify a specific style for each cell
     * @return write result
     */
    public static boolean writeExcelWithOneSheet(File writeFile,
                                                 String sheetName,
                                                 boolean needHead,
                                                 List<? extends BaseRowModel> data,
                                                 ExcelDataHandler handler) throws Exception {
        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        StyleHandler sh = new StyleHandler(dataMap);
        sh.setHandler(handler);
        return writeExcel(writeFile, sheetName, needHead, data, sh);
    }

    /**
     * Write one sheet, Suitable for writing small data volumes
     *
     * @param writeFile Writes data to the specified file
     * @param sheetName sheet name
     * @param needHead  Whether to write header
     * @param data      Write data
     * @return write result
     */
    public static boolean writeExcelWithOneSheet(File writeFile,
                                                 String sheetName,
                                                 boolean needHead,
                                                 List<? extends BaseRowModel> data) throws Exception {
        return writeExcel(writeFile, sheetName, needHead, data, null);
    }

    /**
     * Write multiple sheets, Suitable for writing small data volumes
     *
     * @param writeFile Writes data to the specified file
     * @param data      Write data
     * @param needHead  Whether to write header
     * @param handler   The data processor, through which you can specify a specific style for each cell
     * @return write result
     */
    public static boolean writeExcelWithMultiSheet(File writeFile,
                                                   Map<String, List<? extends BaseRowModel>> data,
                                                   boolean needHead,
                                                   ExcelDataHandler handler) throws Exception {
        StyleHandler sh = new StyleHandler(data);
        sh.setHandler(handler);
        return writeExcelWithHandler(writeFile, data, needHead, sh);
    }

    /**
     * Write multiple sheets, Suitable for writing small data volumes
     *
     * @param writeFile Writes data to the specified file
     * @param data      Write data
     * @param needHead  Whether to write header
     * @return write result
     */
    public static boolean writeExcelWithMultiSheet(File writeFile,
                                                   Map<String, List<? extends BaseRowModel>> data,
                                                   boolean needHead) throws Exception {
        return writeExcelWithHandler(writeFile, data, needHead, null);
    }

    /**
     * Write one sheet
     * We can use ExcelWrapWriter to write data in batches, so that writing large amount of data will not cause OOM
     * @param wrapWriter
     * @param sheetName Sheet name
     * @param needHead Whether to write header
     * @param data
     * @return
     * @throws Exception
     */
    public static boolean writeExcelWithOneSheet(ExcelWrapWriter wrapWriter,
                                                 String sheetName,
                                                 boolean needHead,
                                                 List<? extends BaseRowModel> data) throws Exception {
        if (wrapWriter == null) {
            throw new IllegalArgumentException("Parameter wrapWriter can not be null");
        }

        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        initStyleDataHandler(wrapWriter, dataMap);
        return writeExcelWithMultiSheet(wrapWriter, needHead, dataMap);
    }

    /**
     * Write multi sheet
     * We can use ExcelWrapWriter to write data in batches, so that writing large amount of data will not cause OOM
     * @param wrapWriter
     * @param needHead Whether to write header
     * @param data
     * @return
     * @throws Exception
     */
    public static boolean writeExcelWithMultiSheet(ExcelWrapWriter wrapWriter,
                                                   boolean needHead,
                                                   Map<String, List<? extends BaseRowModel>> data) throws Exception {
        if (wrapWriter == null) {
            throw new IllegalArgumentException("Parameter wrapWriter can not be null");
        }

        if (data == null) {
            throw new IllegalArgumentException("Parameter data can not be null");
        }

        initStyleDataHandler(wrapWriter, data);

        ExcelWriter ew = wrapWriter.getWriter();

        Map<String, Sheet> sheetMap = getSheetMap(wrapWriter, data);
        wrapWriter.setSheetMap(sheetMap);
        for (String sheetName : data.keySet()) {
            List<? extends BaseRowModel> models = data.get(sheetName);
            if (models != null && !models.isEmpty()) {
                models = convertData(models);
            }

            Sheet sheet = sheetMap.get(sheetName);
            ew.write(models, sheet, needHead);
        }
        return true;
    }

    /**
     * init style data
     * @param wrapWriter
     * @param data
     */
    private static void initStyleDataHandler(ExcelWrapWriter wrapWriter, Map<String, List<? extends BaseRowModel>> data) {
        StyleHandler styleHandler = wrapWriter.getStyleHandler();
        if (styleHandler != null) {
            styleHandler.setDataMap(data);
        }
    }

    private static boolean writeExcel(File writeFile,
                                      String sheetName,
                                      boolean needHead,
                                      List<? extends BaseRowModel> data,
                                      WriteHandler handler) throws Exception {
        Map<String, List<? extends BaseRowModel>> dataMap = new HashMap<>(1);
        dataMap.put(sheetName, data);
        return writeExcelWithHandler(writeFile, dataMap, needHead, handler);
    }

    private static Map<String, Sheet> getSheetMap(ExcelWrapWriter wrapWriter, Map<String, List<? extends BaseRowModel>> data) {
        Map<String, Sheet> sheetMap = wrapWriter.getSheetMap();
        if (sheetMap == null) {
            sheetMap = new HashMap<>(data.size());
            int sheetIndex = 1;
            for (String sheetName : data.keySet()) {
                List<? extends BaseRowModel> models = data.get(sheetName);
                Class cls = models.isEmpty() ? BaseRowModel.class : models.get(0).getClass();
                Sheet sheet = new Sheet(sheetIndex, 0, cls);
                sheetName = StringUtils.isEmpty(sheetName) ? "Sheet" + sheetIndex : sheetName;
                sheet.setSheetName(sheetName);
                sheetMap.put(sheetName, sheet);
                sheetIndex++;
            }
        }
        return sheetMap;
    }

    /**
     * Write excel
     *
     * @param writeFile Writes data to the specified file
     * @param data      Write data
     * @param needHead
     * @param handler   @return write result
     */
    private static boolean writeExcelWithHandler(File writeFile,
                                                 Map<String, List<? extends BaseRowModel>> data,
                                                 boolean needHead, WriteHandler handler) throws Exception {
        boolean result = false;
        if (MapUtils.isEmpty(data)) {
            log.warn("write excel data not empty");
            return false;
        }

        ExcelWriter ew = null;
        try {
            ew = EasyExcelFactory.getWriterWithTempAndHandler(null, new FileOutputStream(writeFile),
                    ExcelTypeEnum.XLSX, needHead, handler);

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
            result = true;
            log.info("write excel data success, file path with:{}", writeFile.getPath());
        } catch (Exception e) {
            log.error("write excel data failed , file path with:{}, cause is:{}", writeFile.getPath(), e);

        } finally {
            if (ew != null) {
                ew.finish();
            }

        }
        return result;
    }

    /**
     * Data formatting
     *
     * @param data
     * @throws Exception
     */
    private static <T> List<T> convertData(List<? extends BaseRowModel> data) throws Exception {
        List<T> result = new ArrayList<>();
        for (Object o : data) {
            final Object copyObj = BeanUtils.transform(o.getClass(), o);

            List<Field> fields = new ArrayList<>();
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
     *
     * @param in
     * @param sheet
     * @return read data
     */
    public static List<Object> read(InputStream in, Sheet sheet) {
        return EasyExcelFactory.read(in, sheet);
    }

    /**
     * Parsing large file
     *
     * @param in
     * @param sheet
     * @param listener Callback method after each row is parsed
     */
    public static void readBySax(InputStream in, Sheet sheet, AnalysisEventListener listener) {
        EasyExcelFactory.readBySax(in, sheet, listener);
    }

    /**
     * Get ExcelReader, All the sheet data can be obtained through ExcelReader object
     *
     * @param in
     * @param listener Callback method after each row is parsed
     * @return ExcelReader
     */
    public static ExcelReader getReader(InputStream in, AnalysisEventListener listener) {
        return EasyExcelFactory.getReader(in, listener);
    }
}
