package org.example.service;

import org.example.entity.ReportDto;
import java.util.List;

/**
 * 生成分析报告工具类
 */
public interface BuildReport {

    /**
     * 根据传入参数，生成报告到指定路径
     * @param reportDtos
     * @param filePath
     */
    void build(List<ReportDto> reportDtos, String filePath);

}
