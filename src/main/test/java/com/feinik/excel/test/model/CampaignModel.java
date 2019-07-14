package com.feinik.excel.test.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import com.feinik.excel.annotation.ExcelValueFormat;
import lombok.Data;

import java.io.Serializable;

/**
 * @author Feinik
 */
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

    public CampaignModel() {
    }

    public CampaignModel(String day, String campaignId, String campaignName, String cost, String clicks, String ctr) {
        this.day = day;
        this.campaignId = campaignId;
        this.campaignName = campaignName;
        this.cost = cost;
        this.clicks = clicks;
        this.ctr = ctr;
    }
}
