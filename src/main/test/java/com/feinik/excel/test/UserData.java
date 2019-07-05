package com.feinik.excel.test;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import com.feinik.excel.annotation.ExcelValueFormat;
import lombok.Data;

import java.io.Serializable;

/**
 *
 * @author Feinik
 */

@Data
public class UserData extends BaseRowModel implements Serializable {

    @ExcelProperty(value = "用户名", index = 0)
    private String userName;

    @ExcelProperty(value = "年龄", index = 1)
    private Integer age;

    @ExcelProperty(value = "工资", index = 2)
    @ExcelValueFormat(format = "{0}￥")
    private String salary;

}
