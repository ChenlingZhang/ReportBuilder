package org.example.utils;

import org.example.entity.Analysis;
import org.example.entity.ReportDto;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DataUtils {
    public static Map<String, List<ReportDto>> generateListToMapKey(List<ReportDto> reportDtos, int keyType) {
        Map<String, List<ReportDto>> hashMap = new HashMap<>();
        for (ReportDto reportDto : reportDtos) {
            String generateKey = "";
            switch (keyType) {
                case 1:
                    generateKey = reportDto.getStand();
                    break;
                case 2:
                    generateKey = reportDto.getStand() +"-"+ reportDto.getSystem();
                    break;
                case 3:
                    generateKey = reportDto.getStand() + reportDto.getSystem() + reportDto.getType();
                    break;
            }
            if (hashMap.containsKey(generateKey)) {
                List<ReportDto> currentData = hashMap.get(generateKey);
                currentData.add(reportDto);
                hashMap.put(generateKey, currentData);
            } else {
                List<ReportDto> reportDtoList = new ArrayList<>();
                reportDtoList.add(reportDto);
                hashMap.put(generateKey, reportDtoList);
            }
        }
        return hashMap;
    }

    public static double[] coverArray(ArrayList<Double> list ){
        double[] dl = new double[list.size()];
        for (int i = 0; i < list.size(); i++) {
            dl[i] = list.get(i);
        }
        return dl;
    }
    public static Double getMax(List<Analysis> analysis) {
        double temp = 0;
        for (Analysis a: analysis) {
            if (a.getMax() > temp){
                temp = a.getMax();
            }
        }
        return temp;
    }

    public static Double getMin(List<Analysis> analysis) {
        double temp = 0;
        for (Analysis a: analysis) {
            if (a.getMin() < temp){
                temp = a.getMin();
            }
        }
        return temp;
    }

    public static Double getAvg(List<Analysis> analysis) {
        double total = 0;
        for (int i = 0; i < analysis.size() ; i++) {
            total = total + analysis.get(i).getAvg();
        }
        return  total/analysis.size();
    }
}
