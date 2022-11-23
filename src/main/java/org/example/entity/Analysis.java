package org.example.entity;

import lombok.Data;

import java.time.LocalDateTime;

@Data
public class Analysis {

    /**
     * 日期
     */
    private LocalDateTime date;

    /**
     * 最大值
     */
    private Double max;

    /**
     * 最小值
     */
    private Double min;

    /**
     * 平均值
     */
    private Double avg;

}
