package org.example.service.impl;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphCollection;
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
        try{
            File file = new File(filePath);
            file.createNewFile();
            Document document = new Document(filePath);
            DocumentBuilder builder = new DocumentBuilder(document);
            ParagraphCollection collection = document.getFirstSection().getBody().getParagraphs();
            Map<String,List<ReportDto>> dataMap = DataUtils.generateListToMapKey(reportDtos,2);
            Set<String> keySet = dataMap.keySet();
            log.info("Map生成结束，Key值{}", keySet);
            int mainTitleCount =1;
            for (String key: keySet) {
                DocUtils.generateTitle(builder, mainTitleCount + " " + key, 1);
                List<ReportDto> reports = dataMap.get(key);
                List<String> typeList = new ArrayList<>(); // 二级标题存储表
                List<String> nameList = new ArrayList<>(); // 三级标题存储表
                int typeCount = 0;
                int nameCount = 0;
                int fixTitlteCount = 0;
                for (ReportDto report: reports) {

                    if (!typeList.contains(report.getType())){
                        typeCount++;
                        String title = mainTitleCount+"."+typeCount+" " +report.getType();

                        DocUtils.generateTitle(builder,title,2);
                        typeList.add(report.getType());

                    }
                    if (!nameList.contains(report.getName())){
                        nameCount++;
                        String title = mainTitleCount+"."+typeCount+"."+nameCount + " "+ report.getName();
                        DocUtils.generateTitle(builder,title,3);
                        nameList.add(report.getName());

                    }
                    fixTitlteCount++;
                    String fixTitle = mainTitleCount+"."+typeCount+"."+nameCount+"."+fixTitlteCount+"统计通道"+" ";
                    DocUtils.generateTitle(builder,fixTitle,4);
                    log.info("key：{} title生成结束",key);

                    log.info("key{}: 开始生成通道信息表",key);
                    DocUtils.textCenterShow(builder);
                    builder.writeln("通道表"+channelTableCount+":"+report.getStand()+report.getSystem()+report.getType());
                    DocUtils.createChannelTable(reports,builder,document);
                    log.info("key{}: 开始生数据表",key);
                    DocUtils.createDataTable(reports,builder);
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
