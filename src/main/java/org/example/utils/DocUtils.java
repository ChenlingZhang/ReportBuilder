package org.example.utils;

import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.Shape;
import lombok.extern.slf4j.Slf4j;
import org.example.entity.Analysis;
import org.example.entity.ReportDto;

import java.awt.*;
import java.io.File;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;


@Slf4j
public class DocUtils {
    /**
     * 用于生成文章1-4级的标题
     * @param builder DocumentBuilder
     * @param title 需要生成的文章标题
     * @param titleLevel 需要生成的标题字号1-4
     * @throws Exception 统一错误抛出
     */
    public static void generateTitle(DocumentBuilder builder, String title, int titleLevel) throws Exception {
        switch (titleLevel){
            case 1:
                builder.insertHtml("<h1 style='text-align:left;font-family:Simsun;'>" + title + "</h1>");
                log.info("生成1级标题{}", title);
                break;
            case 2:
                builder.insertHtml("<h2 style='text-align:left;font-family:Simsun;'>" + title + "</h2>");
                log.info("生成2级标题{}", title);
                break;
            case 3:
                builder.insertHtml("<h3 style='text-align:left;font-family:Simsun;'>" + title + "</h3>");
                log.info("生成3级标题{}", title);
                break;
            case 4:
                builder.insertHtml("<h4 style='text-align:left;font-family:Simsun;'>" + title + "</h4>");
                log.info("生成4级标题{}", title);
                break;

        }
    }
    /**
     * 生成文件
     * @param path 传入路径+文件名例如 /doc/test.docx
     * @return
     * @throws Exception
     */
    public static File createFile(String path) throws Exception {
        File file = new File(path);
        if (file.exists()) {
            throw new Exception("文件创建失败，路径中已存在改文件或传入的不是一个文件");
        }
        file.createNewFile();
        return file;
    }

    /**
     * 设置字体颜色
     * @param builder documentBuilder
     * @param color font color
     * @return builder
     */
    public static DocumentBuilder setFontColor(DocumentBuilder builder, Color color){
        Font font = builder.getFont();
        font.setColor(color);
        return builder;
    }

    /**
     * 居中显示文字
     * @param builder documentbuilder
     */
    public static void textCenterShow(DocumentBuilder builder){
        ParagraphFormat format = builder.getParagraphFormat();
        format.setAlignment(ParagraphAlignment.CENTER);
    }
    /**
     * 靠左显示文字
     * @param builder documentbuilder
     */
    public static void textLeftShow(DocumentBuilder builder){
        ParagraphFormat format = builder.getParagraphFormat();
        format.setAlignment(ParagraphAlignment.LEFT);
    }

    /**
     * 获取时间格式
     * @param dateTime
     * @return
     */
    public static String dateFormat(LocalDateTime dateTime){
        return dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
    }

    /**
     * 获取日期
     * @param dateTime
     * @return
     */
    public static int getDate(LocalDateTime dateTime){
        return dateTime.getDayOfMonth();
    }

    public static void createChannelTable(List<ReportDto> reportDtos, DocumentBuilder builder, Document document){
        int count = 0;
        Table table = builder.startTable();
        Row row = new Row(document);
        row.getCells().add(createCell("对象名称",document));
        row.getCells().add(createCell("通道号",document));
        row.getCells().add(createCell("单位",document));
        row.getCells().add(createCell("要求",document));
        table.getRows().insert(count,row);
        count++;
        count++;
        for (ReportDto reportDto : reportDtos)  {
            table.getRows().insert(count,addRowsBehind(reportDto,document));
        }


        builder.endTable();
    }

    public static void createDataTable(List<ReportDto> reportDtos, DocumentBuilder builder){
        int count = 0;
        for (int i = 0; i < reportDtos.size() ; i++) {
            count++;
            tableDesc(builder,reportDtos.get(i),count);
            builder.startTable();
            builder.insertCell();
            builder.getCellFormat().setWidth(20);
            cellCenterShow(builder);
            builder.write("参数");
            combineCells(builder,3,reportDtos.get(i).getChannel());
            builder.endRow();
            builder.insertCell();
            builder.getCellFormat().setWidth(20);
            builder.write("要求");
            combineCells(builder,3,reportDtos.get(i).getLowSet()+"~"+reportDtos.get(i).getHighSet());
            builder.endRow();
            String[] context = {"日期","max","min","avg"};
            createDataTableSeconLine(context,builder);
            builder.endRow();
            for (Analysis analysis: reportDtos.get(i).getAnalysis()) {
                insertData(reportDtos.get(i),analysis,builder);
                builder.endRow();
            }
            builder.endTable();
            log.info("开始生成数据图");
            createChannelGraphic(reportDtos.get(i),builder,count);
        }
    }
    private static void createChannelGraphic(ReportDto reportDto, DocumentBuilder builder, int index){
        ArrayList<String> time = new ArrayList<>();
        ArrayList<Double> maxArray = new ArrayList<>();
        ArrayList<Double> minArray = new ArrayList<>();
        ArrayList<Double> avgArray = new ArrayList<>();

        builder.getParagraphFormat().setFirstLineIndent(0);
        builder.writeln("");
        try {
            Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
            shape.setVerticalAlignment(VerticalAlignment.CENTER);
            shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
            shape.setAllowOverlap(false);
            Chart chart = shape.getChart();
            ChartTitle title = chart.getTitle();
            String graphicTitle = "图" + index +":" + reportDto.getStand()+reportDto.getSystem()+reportDto.getType()+reportDto.getChannel()+" 数据统计图";
            title.setText(graphicTitle);
            // 获取案例中所有的Series并清空
            ChartSeriesCollection seriesCollections = chart.getSeries();
            seriesCollections.clear();

            for (Analysis analysis : reportDto.getAnalysis()) {
                time.add(dateFormat(analysis.getDate()));
                maxArray.add(analysis.getMax());
                minArray.add(analysis.getMin());
                avgArray.add(analysis.getAvg());
            }
            String[] times = time.toArray(new String[0]);
            double[] max = DataUtils.coverArray(maxArray);
            double[] min = DataUtils.coverArray(minArray);
            double[] avg = DataUtils.coverArray(avgArray);
            double[] lowSet = new double[times.length];
            double[] highSet = new double[times.length];
            for (int i = 0; i < lowSet.length; i++) {
                lowSet[i] = reportDto.getLowSet();
            }
            for (int j = 0; j < highSet.length; j++) {
                lowSet[j] = reportDto.getHighSet();
            }


            seriesCollections.add("Max", times, max);
            seriesCollections.add("min", times, min);
            seriesCollections.add("avg", times, avg);
            seriesCollections.add("标定最大值", times, lowSet);
            seriesCollections.add("标定最小值", times, highSet);
            ChartSeries minSeries = seriesCollections.get(3);
            minSeries.getMarker().setSymbol(MarkerSymbol.CIRCLE);
            ChartSeries maxSeries = seriesCollections.get(4);
            maxSeries.getMarker().setSymbol(MarkerSymbol.CIRCLE);
            builder.writeln("");
            builder.getParagraphFormat().setFirstLineIndent(2);
            builder.writeln("");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    // 单元格内文字居中显示
    private static void cellCenterShow(DocumentBuilder builder){
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
    }
    private static void insertData(ReportDto reportDto, Analysis analysis, DocumentBuilder builder) {
        builder.insertCell();
        setFontColor(builder, Color.BLACK);
        builder.getCellFormat().setWidth(20);
        builder.getCellFormat().setVerticalMerge(CellMerge.NONE);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        cellCenterShow(builder);
        builder.write(Integer.toString(getDate(analysis.getDate())));

        builder.insertCell();
        setFontColor(builder, Color.BLACK);
        builder.getCellFormat().setWidth(20);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        cellCenterShow(builder);
        if (analysis.getMax() > reportDto.getHighSet() ){
            setFontColor(builder, Color.RED);
        }
        builder.write(analysis.getMax().toString());
        setFontColor(builder, Color.BLACK);

        builder.insertCell();
        builder = setFontColor(builder, Color.BLACK);
        builder.getCellFormat().setWidth(20);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        cellCenterShow(builder);
        if (analysis.getMin() < reportDto.getLowSet() ){
            setFontColor(builder, Color.RED);
        }
        builder.write(analysis.getMin().toString());
        setFontColor(builder, Color.BLACK);

        builder.insertCell();
        builder = setFontColor(builder, Color.BLACK);
        builder.getCellFormat().setWidth(20);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        cellCenterShow(builder);
        if (analysis.getAvg() > reportDto.getHighSet() || analysis.getAvg() < reportDto.getLowSet() ){
            setFontColor(builder, Color.RED);
        }
        builder.write(analysis.getAvg().toString());
        setFontColor(builder, Color.BLACK);
    }

    private static void createDataTableSeconLine(String[] context, DocumentBuilder builder) {
        for (int i = 0; i < context.length ; i++) {
            builder.insertCell();
            cellCenterShow(builder);
            builder.getCellFormat().setHorizontalMerge(CellMerge.NONE);
            builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
            builder.write(context[i]);
        }
    }


    /**
     * 创建通道表数据栏
     * @param reportDto
     * @param document
     * @return
     */
    private static Row addRowsBehind(ReportDto reportDto, Document document) {

        Row data = new Row(document);
        data.getCells().add(createCell(reportDto.getName(),document));
        data.getCells().add(createCell(reportDto.getChannel(),document));
        data.getCells().add(createCell(reportDto.getSpecification(),document));
        data.getCells().add(createCell(reportDto.getLowSet()+"~"+reportDto.getHighSet(),document));

        return data;
    }

    /**
     * 创建单元格
     * @param value 单元格内容
     * @param document document
     * @return
     */
    private static Cell createCell(String value, Document document) {
        Cell cell = new Cell(document);
        Paragraph p = new Paragraph(document);
        p.appendChild(new Run(document,value));
        cell.appendChild(p);
        return cell;
    }

    /**
     * 生成数据表描述
     * @param builder builder
     * @param report dto
     * @param count datatablecount
     */
    private static void tableDesc(DocumentBuilder builder, ReportDto report,int count){
        String month = Integer.toString(report.getAnalysis().get(0).getDate().getMonthValue());
        String stand = report.getStand();
        String systemName = report.getSystem();
        String type = report.getType();
        String device = report.getDevice();
        String channelName = report.getChannel();
        Double max = DataUtils.getMax(report.getAnalysis());
        Double min = DataUtils.getMin(report.getAnalysis());
        Double avg = DataUtils.getAvg(report.getAnalysis());
        textLeftShow(builder);
        builder.write(month + "月1日至31日" + stand+systemName+type+channelName+"变化范围为:");
        if (min < report.getLowSet()){
            setFontColor(builder, Color.RED);
        }
        builder.write(min.toString());
        setFontColor(builder, Color.BLACK);
        builder.write("~");
        if (max > report.getHighSet()){
            setFontColor(builder, Color.RED);
        }
        builder.write(max.toString());
        builder = setFontColor(builder,Color.BLACK);;
        builder.write("平均值为：");
        if (avg < report.getLowSet()||avg<report.getHighSet()){
            setFontColor(builder, Color.RED);
        }
        builder.write(String.format("%.3f",avg));
        setFontColor(builder, Color.BLACK);
        builder.writeln(" ");
        textCenterShow(builder);
        builder.writeln("数据表"+count+":"+stand+systemName+type+channelName+"变化表");
        textLeftShow(builder);
    }
    private static void combineCells(DocumentBuilder builder, int index, String context){
        builder.insertCell();
        builder.getCellFormat().setWidth(20);
        builder.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
        builder.write(context);
        for (int i = 0; i < index-1 ; i++) {
            builder.insertCell();
            builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
        }

    }
}
