import lombok.extern.slf4j.Slf4j;
import org.example.entity.Analysis;
import org.example.entity.ReportDto;
import org.example.service.impl.BuildReportImpl;
import org.junit.Test;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class MethodTest {
    private static List<ReportDto> createData(){
        List<ReportDto> reportDtos = new ArrayList<>();
        List<Analysis> analyses = new ArrayList<>();
        Analysis analysis1 = new Analysis();
        LocalDateTime time = LocalDateTime.now();
        analysis1.setDate(time);
        analysis1.setMin(5.0);
        analysis1.setMax(8.0);
        analysis1.setAvg(6.0);

        Analysis analysis2 = new Analysis();
        analysis2.setDate(time);
        analysis2.setMin(5.0);
        analysis2.setMax(6.0);
        analysis2.setAvg(5.5);

        Analysis analysis3 = new Analysis();
        analysis3.setDate(time);
        analysis3.setMin(10.0);
        analysis3.setMax(20.0);
        analysis3.setAvg(15.0);

        analyses.add(analysis1);
        analyses.add(analysis2);
        analyses.add(analysis3);

        ReportDto reportDto1 = new ReportDto();
        reportDto1.setStand("516");
        reportDto1.setSystem("安全环境监测系统");
        reportDto1.setType("温度变送器");
        reportDto1.setName("name1");
        reportDto1.setSpecification("sp1");
        reportDto1.setChannel("channel1");
        reportDto1.setLowSet(2.0);
        reportDto1.setHighSet(10.0);
        reportDto1.setAnalysis(analyses);

        ReportDto reportDto2 = new ReportDto();
        reportDto2.setStand("516");
        reportDto2.setSystem("xxxx系统");
        reportDto2.setType("xxxx类型");
        reportDto2.setName("name2");
        reportDto2.setSpecification("sp2");
        reportDto2.setChannel("channel2");
        reportDto2.setLowSet(2.0);
        reportDto2.setHighSet(10.0);
        reportDto2.setAnalysis(analyses);

        ReportDto reportDto5 = new ReportDto();
        reportDto5.setStand("516");
        reportDto5.setSystem("xxxx系统");
        reportDto5.setType("xxxx类型");
        reportDto5.setName("name2");
        reportDto5.setSpecification("sp2");
        reportDto5.setChannel("channel2");
        reportDto5.setLowSet(2.0);
        reportDto5.setHighSet(10.0);
        reportDto5.setAnalysis(analyses);
        ReportDto reportDto51 = new ReportDto();
        reportDto51.setStand("516");
        reportDto51.setSystem("xxxx系统");
        reportDto51.setType("xxxx类型");
        reportDto51.setName("name2");
        reportDto51.setSpecification("sp2");
        reportDto51.setChannel("channel2");
        reportDto51.setLowSet(2.0);
        reportDto51.setHighSet(10.0);
        reportDto51.setAnalysis(analyses);
        ReportDto reportDto52 = new ReportDto();
        reportDto52.setStand("516");
        reportDto52.setSystem("xxxx系统");
        reportDto52.setType("xxxx类型");
        reportDto52.setName("name2");
        reportDto52.setSpecification("sp2");
        reportDto52.setChannel("channel2");
        reportDto52.setLowSet(2.0);
        reportDto52.setHighSet(10.0);
        reportDto52.setAnalysis(analyses);
        ReportDto reportDto53 = new ReportDto();
        reportDto53.setStand("516");
        reportDto53.setSystem("xxxx系统");
        reportDto53.setType("xxxx类型");
        reportDto53.setName("name2");
        reportDto53.setSpecification("sp2");
        reportDto53.setChannel("channel2");
        reportDto53.setLowSet(2.0);
        reportDto53.setHighSet(10.0);
        reportDto53.setAnalysis(analyses);
        ReportDto reportDto54 = new ReportDto();
        reportDto54.setStand("516");
        reportDto54.setSystem("xxxx系统");
        reportDto54.setType("xxxx类型");
        reportDto54.setName("name2");
        reportDto54.setSpecification("sp2");
        reportDto54.setChannel("channel2");
        reportDto54.setLowSet(2.0);
        reportDto54.setHighSet(10.0);
        reportDto54.setAnalysis(analyses);

        ReportDto reportDto4 = new ReportDto();
        reportDto4.setStand("516");
        reportDto4.setSystem("xxx系统");
        reportDto4.setType("xxxx类型");
        reportDto4.setName("name2");
        reportDto4.setSpecification("sp2");
        reportDto4.setChannel("channel2");
        reportDto4.setLowSet(2.0);
        reportDto4.setHighSet(10.0);
        reportDto4.setAnalysis(analyses);

        ReportDto reportDto3 = new ReportDto();
        reportDto3.setStand("517");
        reportDto3.setSystem("xxxx系统");
        reportDto3.setType("xxxx类型");
        reportDto3.setName("name2");
        reportDto3.setSpecification("sp2");
        reportDto3.setChannel("channel2");
        reportDto3.setLowSet(2.0);
        reportDto3.setHighSet(10.0);
        reportDto3.setAnalysis(analyses);

        reportDtos.add(reportDto1);
        reportDtos.add(reportDto2);
        reportDtos.add(reportDto3);
        reportDtos.add(reportDto4);
        reportDtos.add(reportDto5);
        reportDtos.add(reportDto51);
        reportDtos.add(reportDto52);
        reportDtos.add(reportDto53);
        reportDtos.add(reportDto54);

        return reportDtos;
    }

    @Test
    public void test(){
        BuildReportImpl builder = new BuildReportImpl();
        builder.build(createData(), "./doc/test.docx");
    }

    @Test
    public void commonTest(){
        log.info(""+(12%4));
    }
}
