package com.feinik.excel.test.util;

import java.io.InputStream;

/**
 *
 * @author Feinik
 */
public class FileUtil {

    public static InputStream getResourcesFileInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }
}
