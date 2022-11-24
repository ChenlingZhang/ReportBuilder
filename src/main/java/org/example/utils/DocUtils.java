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

    public static Table createChannelTable(DocumentBuilder builder, Document document){
        Table table = builder.startTable();
        Row row = new Row(document);
        row.getCells().add(createCell("对象名称",document));
        row.getCells().add(createCell("通道号",document));
        row.getCells().add(createCell("单位",document));
        row.getCells().add(createCell("要求",document));
        table.getRows().add(row);
        builder.endTable();
        return table;
    }

    public static Table initDataTable(DocumentBuilder builder,Document document, ReportDto report){
        Table table = builder.startTable();
        List<Analysis> analysises = report.getAnalysis();
        RowCollection rows= table.getRows();
        Row paraRow = new Row(document);
        paraRow.getCells().add(createCell("参数",document));
        Row requirementsRow = new Row(document);
        requirementsRow.getCells().add(createCell("要求",document));
        Row dateRow = new Row(document);
        dateRow.getCells().add(createCell("日期",document));
        table.appendChild(paraRow);
        table.appendChild(requirementsRow);
        table.appendChild(dateRow);
        analysises.forEach(analysis -> {
            Row row = new Row(document);
            row.getCells().add(createCell(getDate(analysis.getDate())+"",document));
            table.appendChild(row);
        });
        return table;
    }
    public static int insertIntoDataTable(Document document, DocumentBuilder builder, ReportDto report, Table table,int ec)  {
        RowCollection rows = table.getRows();
        int startPoint = rows.get(0).getCount(); //获取插入前第一行中元素的数量
        int secondPoint = startPoint +1;
        int thirdPoint = secondPoint + 1;
        int executeCount =ec;

        // 第一行插入固定数据 channelName 并合并3个单元格
        Row row1 = rows.get(0);
        Cell channelCell = createCell(report.getChannel(),document);
        channelCell.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
        cellCenterShow(builder);
        row1.getCells().insert(startPoint, channelCell);
        combineCells(row1,document,secondPoint,thirdPoint);

        // 第二行插入固定要求数据
        Row row2 = rows.get(1);
        Cell requireCell = createCell(report.getLowSet()+"~"+report.getHighSet(),document);
        requireCell.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
        row2.getCells().insert(startPoint,requireCell);
        combineCells(row2,document,secondPoint,thirdPoint);

        // 第三行内容插入
        Row row3 = rows.get(2);
        Cell maxCell = createCell("max",document);
        row3.getCells().insert(startPoint,maxCell);
        Cell minCell = createCell("min",document);
        row3.getCells().insert(secondPoint,minCell);
        Cell avgCell = createCell("avg",document);
        row3.getCells().insert(thirdPoint,avgCell);

        // 数据插入
          for(int j = 3; j < rows.getCount(); j++) {
              Analysis a = report.getAnalysis().get(j-3);
              Row dataRow = table.getRows().get(j);
              log.info(j+"");
              double max = DataUtils.getMax(report.getAnalysis());
              double min = DataUtils.getMin(report.getAnalysis());;
              double avg = DataUtils.getAvg(report.getAnalysis());;
              Cell maxDataCell;
              Cell minDataCell;
              Cell avgDataCell;
              if (max>report.getHighSet()){
                  maxDataCell = createRedCellElement(max + "", Color.RED, document);
              }
              else {
                  maxDataCell = createCell(max + "", document);
              }
              if (min<report.getLowSet()){
                  minDataCell = createRedCellElement(min+"",Color.RED,document);
              }
              else {
                  minDataCell = createCell(min+"",document);
              }
              if(avg<report.getLowSet() || avg >report.getHighSet()){
                  avgDataCell = createRedCellElement(avg+"",Color.RED,document);
              }
              else {
                  avgDataCell = createCell(String.format("%.3f",avg),document);
              }
              dataRow.getCells().add(maxDataCell);
              dataRow.getCells().add(minDataCell);
              dataRow.getCells().add(avgDataCell);
          }

          return executeCount +1;
  }

    public static Shape initGraphic(ReportDto reportDto, DocumentBuilder builder, int count) throws Exception {
        builder.getParagraphFormat().setFirstLineIndent(0);
        builder.writeln("");
        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
        shape.setVerticalAlignment(VerticalAlignment.CENTER);
        shape.setHorizontalAlignment(HorizontalAlignment.CENTER);
        shape.setAllowOverlap(false);
        Chart chart = shape.getChart();
        ChartTitle chartTitle = chart.getTitle();
        String titleTest = "图" + count + "\t" + generateGraphicTitle(reportDto);
        chartTitle.setText(titleTest);
        // 获取案例中所有的Series并清空
        ChartSeriesCollection seriesCollections = chart.getSeries();
        seriesCollections.clear();
        return shape;
    }

    public static void addLineToGraphic(Shape shape, DocumentBuilder builder, ReportDto reportDto){
        Chart chart = shape.getChart();
        ChartSeriesCollection series = chart.getSeries();
        List<String> time = new ArrayList<>();
        List<Double> maxList = new ArrayList<>();
        List<Double> minList = new ArrayList<>();
        List<Double> avgList = new ArrayList<>();
        reportDto.getAnalysis().forEach(analysis -> {
            time.add(analysis.getDate().toString());
            maxList.add(Double.valueOf(analysis.getMax().toString()));
            minList.add(Double.valueOf(analysis.getMin().toString()));
            avgList.add(Double.valueOf(analysis.getAvg().toString()));
        });
        String[] times = time.toArray(new String[0]);
        double[] max =DataUtils.coverArray(maxList);
        double[] min =DataUtils.coverArray(minList);
        double[] avg =DataUtils.coverArray(avgList);
        double[] lowSet = new double[times.length];
        double[] highSet = new double[times.length];
        for (int i = 0; i < lowSet.length; i++) {
            lowSet[i] = reportDto.getLowSet();
        }
        for (int j = 0; j < highSet.length; j++) {
            lowSet[j] = reportDto.getHighSet();
        }


        series.add("Max", times, max);
        series.add("min", times, min);
        series.add("avg", times, avg);
        series.add("标定最大值", times, lowSet);
        series.add("标定最小值", times, highSet);
        ChartSeries minSeries = series.get(3);
        minSeries.getMarker().setSymbol(MarkerSymbol.CIRCLE);
        ChartSeries maxSeries = series.get(4);
        maxSeries.getMarker().setSymbol(MarkerSymbol.CIRCLE);
        builder.writeln("");
        builder.getParagraphFormat().setFirstLineIndent(2);
        builder.writeln("");
    }
    private static String generateGraphicTitle(ReportDto reportDto){
        String channelName = reportDto.getChannel();
        int month = reportDto.getAnalysis().get(0).getDate().getMonthValue();
        int day = reportDto.getAnalysis().get(0).getDate().getDayOfMonth();
        String startDate = month+"月"+day+"日";
        int endMonth = reportDto.getAnalysis().get(reportDto.getAnalysis().size()-1).getDate().getMonthValue();
        int endDay = reportDto.getAnalysis().get(reportDto.getAnalysis().size()-1).getDate().getDayOfMonth();
        String endDate = endMonth+"月"+endDay+"日";
        return startDate+" 至 "+ endDate + channelName + " 数据统计图";
    }
    // 单元格内文字居中显示
    private static void cellCenterShow(DocumentBuilder builder){
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
    }

    /**
     * 创建通道表数据栏
     * @param reportDto
     * @param document
     * @return
     */
    public static void addRowsBehind(Table table,ReportDto reportDto, Document document) {

        Row data = new Row(document);
        data.getCells().add(createCell(reportDto.getName(),document));
        data.getCells().add(createCell(reportDto.getChannel(),document));
        data.getCells().add(createCell(reportDto.getSpecification(),document));
        data.getCells().add(createCell(reportDto.getLowSet()+"~"+reportDto.getHighSet(),document));
        table.getRows().add(data);
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
        p.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        p.appendChild(new Run(document,value));
        cell.getCellFormat().setWidth(20);
        cell.appendChild(p);
        return cell;
    }

    private static Cell createRedCellElement(String value, Color color, Document document){
        Paragraph p = new Paragraph(document);
        Cell cell = new Cell(document);
        p.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        Run run = new Run(document,value);
        run.getFont().setColor(color);
        p.appendChild(run);
        cell.getCellFormat().setWidth(20);
        cell.appendChild(p);
        return cell;
    }

    public static void tableDesc(DocumentBuilder builder, Document document, ReportDto report, Paragraph paragraph){
        String month = Integer.toString(report.getAnalysis().get(0).getDate().getMonthValue());
        String stand = report.getStand();
        String systemName = report.getSystem();
        String type = report.getType();
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
    }
    private static void combineCells(Row row, Document document, int index1, int index2){
        Cell commonEmptyCell1 = createCell("",document);
        Cell commonEmptyCell2 = createCell("",document);

        commonEmptyCell1.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
        row.getCells().insert(index1, commonEmptyCell1);
        commonEmptyCell2.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
        row.getCells().insert(index2, commonEmptyCell2);
    }
}
