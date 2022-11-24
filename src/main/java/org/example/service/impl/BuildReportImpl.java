package org.example.service.impl;

import com.aspose.words.*;
import lombok.extern.slf4j.Slf4j;
import org.example.entity.ReportDto;
import org.example.service.BuildReport;
import org.example.utils.DataUtils;
import org.example.utils.DocUtils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Slf4j
public class BuildReportImpl implements BuildReport {
    /**
     * 根据传入参数，生成报告到指定路径
     *
     * @param reportDtos
     * @param filePath
     */
    @Override
    public void build(List<ReportDto> reportDtos, String filePath) {
        int channelTableCount = 1;
        int dataTableCount =1;
        int shapeCount = 1;
        try{
            File file = new File(filePath);
            file.createNewFile();
            Document document = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(document);

            // 将纸张设置为横向排列
            log.info("纸张设置为横向排列");
            SectionCollection sectionCollection = document.getSections();
            sectionCollection.forEach(section->{section.getPageSetup().setOrientation(Orientation.LANDSCAPE);});

            Map<String,List<ReportDto>> dataMap = DataUtils.generateListToMapKey(reportDtos,2);
            Set<String> keySet = dataMap.keySet();
            log.info("Map生成结束，Key值{}", keySet);
            int mainTitleCount =1;

            for (String key: keySet) {
                DocUtils.generateTitle(builder, mainTitleCount + " " + key, 1);
                List<ReportDto> reports = dataMap.get(key);
                List<String> typeList = new ArrayList<>(); // 二级标题存储表
                int typeCount = 0;
                int titleCount = 0;
                int fixTitleDisplayCount = 0;
                int index = 0;
                Table channelTable = null;
                Table dataTable = null;
                Paragraph dataDesc = null;
                Shape shape = null;
                int excuteCount=0;
                for (ReportDto report: reports) {
                    int size = reports.size();
                    if (!typeList.contains(report.getType())){
                        typeCount++;
                        String title = mainTitleCount+"."+typeCount+" " +report.getType();

                        DocUtils.generateTitle(builder,title,2);
                        typeList.add(report.getType());
                    }
                    titleCount++;
                    if (fixTitleDisplayCount == 0){
                        String fixTitle = mainTitleCount+"."+typeCount+ "."+titleCount+"统计通道"+" ";
                        DocUtils.generateTitle(builder,fixTitle,3);
                        fixTitleDisplayCount++;
                    }
                    log.info("key：{} title生成结束",key);
                    if (channelTable == null){
                        log.info("channelTable 为null, key{}: 开始生成通道信息表",key);
                        DocUtils.textCenterShow(builder);
                        builder.writeln("通道表"+channelTableCount+":"+report.getStand()+report.getSystem()+report.getType());
                        channelTable = DocUtils.createChannelTable(builder,document);
                        channelTableCount++;
                    }
                    DocUtils.addRowsBehind(channelTable,report,document);
                    builder.writeln("");
                    if (dataTable == null ){
                        log.info("dataTable 为null, key{}: 开始生成数据表",key);
                        DocUtils.textCenterShow(builder);
                        builder.writeln("数据表"+dataTableCount+":"+report.getStand()+report.getSystem()+report.getType());
                        dataTable = DocUtils.initDataTable(builder,document,report);
                        dataTableCount++;
                        excuteCount =0;
                    }
                    log.info("dataTable 不为null, key{}: 开始插入数据",key);
                    excuteCount = DocUtils.insertIntoDataTable(document,builder,report,dataTable,excuteCount);
                    log.info("执行次数{}",excuteCount);

                    if (shape == null){
                        log.info("shape为空开始创建shape模板并清楚模板内数据");
                        shape = DocUtils.initGraphic(report,builder,shapeCount);
                        shapeCount++;
                    }
                    DocUtils.addLineToGraphic(shape,builder,report);
                }
                mainTitleCount++;
            }

            document.save(filePath);
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
}
