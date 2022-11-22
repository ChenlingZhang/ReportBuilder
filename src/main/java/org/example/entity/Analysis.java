package org.example.entity;

import lombok.Data;

@Data
public class Analysis {

    /**
     * 日期
     */
    private Long date;

    /**
     * 最大值
     */
    private String max;

    /**
     * 最小值
     */
    private String min;

    /**
     * 平均值
     */
    private String avg;
}
