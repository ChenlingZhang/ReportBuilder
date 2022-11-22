package org.example.service.impl;

import org.example.entity.ReportDto;
import org.example.service.BuildReport;

import java.util.List;

public class BuildReportImpl implements BuildReport {
    /**
     * 根据传入参数，生成报告到指定路径
     *
     * @param reportDtos
     * @param filePath
     */
    @Override
    public void build(List<ReportDto> reportDtos, String filePath) {

    }
}
