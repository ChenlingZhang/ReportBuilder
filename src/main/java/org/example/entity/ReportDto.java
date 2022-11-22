package org.example.entity;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

/**
 * 分析报告构造 dto
 */
@Data
public class ReportDto {

    /**
     * 工位（厂房）
     * 对应报告：一级标题前缀
     */
    private String stand;

    /**
     * 系统（子系统）
     * 对应报告：一级标题后缀
     */
    private String system;

    /**
     * 设备类型（对象类型）
     * 对应报告：二级标题
     */
    private String type;

    /**
     * 设备编号
     * 没用到，保留
     */
    private String device;

    /**
     * 设备名称（对象名称）
     * 对应报告：参数名称
     */
    private String name;

    /**
     * 设备单位（对象单位）
     * 对应报告：单位
     */
    private String specification;

    /**
     * 通道编号
     * 对应报告：代号
     */
    private String channel;

    /**
     * 报警上限
     * 对应报告：要求
     */
    private String highSet;

    /**
     * 报警下限
     * 对应报告：要求
     */
    private String lowSet;

    /**
     * 分析结果
     */
    private List<Analysis> analysis;

}

