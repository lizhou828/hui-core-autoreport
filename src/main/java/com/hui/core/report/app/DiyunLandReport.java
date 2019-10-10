package com.hui.core.report.app;
import	java.text.SimpleDateFormat;

import com.hui.core.report.model.*;
import com.hui.core.report.util.PowerPointGenerator;

import java.io.File;
import java.util.*;


/**
 * <b><code>ReportApp</code></b>
 * <p/>
 * Description:
 * <p/>
 * <b>Creation Time:</b> 2018/10/31 22:48.
 *
 * @author Hu Weihui
 */
public class DiyunLandReport {

    public static void main(String[] args) {
        String templateFile = System.getProperty("user.dir") +
                File.separator + "src\\main\\resources\\diyun_land_report_template.pptx";

        String resultFile = System.getProperty("user.dir") +
                File.separator + "src\\main\\resources\\diyun_land_report_result.pptx";

        Map<Integer, SlideData> data = getData();

        PowerPointGenerator.generatorPowerPoint(templateFile, resultFile, data);
    }

    /**
     * 造点数据玩完吧
     *
     * @return
     */
    private static Map<Integer, SlideData> getData() {
        Map<Integer, SlideData> map = new HashMap<>();
        //第1页text
        SlideData slideData3 = new SlideData();
        Map<String, String> textDataTest = getTextDataTest();
        slideData3.setTextMap(textDataTest);
        map.put(1, slideData3);


        //第4页表格
        SlideData slideData4 = new SlideData();
        slideData4.setTableDataList(getTableTest1());

        //第4页文本
        slideData4.setTextMap( getTextDataTest2());

        //第4页柱状图
//        slideData4.setChartDataList( getChartData2()); //插入图标的 数据后，打开后，提示“内容有问题，尝试修复”，修复后，就丢失了图标内容


        map.put(4, slideData4);
        return map;
    }

    private static Map<String, String> getTextDataTest() {
        Map<String, String> textMap = new HashMap<>();
        String landLocation = "惠州市江北东区JBD81-01-02-01地块";
        if(!landLocation.contains("地块") && !landLocation.contains("地段")  && !landLocation.contains("片区") && !landLocation.contains("用地")){
            landLocation += "地块";
        }
        textMap.put("landLocation", landLocation);
        SimpleDateFormat yearSDF = new SimpleDateFormat("yyyy");
        String year = yearSDF.format(new Date());
        SimpleDateFormat monthSDF = new SimpleDateFormat("MM");
        String month = monthSDF.format(new Date());
        textMap.put("year", year);
        textMap.put("month", month);
        return textMap;
    }

    private static Map<String, String> getTextDataTest2() {
        Map<String, String> textMap = new HashMap<>();
        textMap.put("year", "2019");
        textMap.put("cityName", "广东省深圳市");

        return textMap;
    }

    private static List<TableData> getTableTest1() {
        TableRowData tableRowData1 = new TableRowData();
        List<String> strings11 = new ArrayList<>();
        strings11.add("地区生产总值(万元)");
        strings11.add("12");
        strings11.add("13");
        strings11.add("14");
        strings11.add("15");
        strings11.add("16");
        tableRowData1.setDataList(strings11);
        TableRowData tableRowData2 = new TableRowData();
        List<String> strings12 = new ArrayList<>();
        strings12.add("增长速度(%)");
        strings12.add("2.2");
        strings12.add("2.3");
        strings12.add("2.4");
        strings12.add("2.5");
        strings12.add("2.6");
        tableRowData2.setDataList(strings12);

        List<TableRowData> tableRowDataList = new ArrayList<>();
        tableRowDataList.add(tableRowData1);
        tableRowDataList.add(tableRowData2);

        TableData tableData = new TableData();
        tableData.setTableRowDataList(tableRowDataList);

        List<TableData> list = new ArrayList<>();
        list.add(tableData);


        TableRowData tableRowData3 = new TableRowData();
        strings11 = new ArrayList<>();
        strings11.add("人均生产总值(万元)");
        strings11.add("2");
        strings11.add("3");
        strings11.add("4");
        strings11.add("5");
        strings11.add("6");
        tableRowData3.setDataList(strings11);
        TableRowData  tableRowData4 = new TableRowData();
        strings12 = new ArrayList<>();
        strings12.add("增长速度(%)");
        strings12.add("0.2");
        strings12.add("0.3");
        strings12.add("0.4");
        strings12.add("0.5");
        strings12.add("0.6");
        tableRowData4.setDataList(strings12);

        List<TableRowData> tableRowDataList2 = new ArrayList<>();
        tableRowDataList2.add(tableRowData3);
        tableRowDataList2.add(tableRowData4);

        TableData tableData2 = new TableData();
        tableData2.setTableRowDataList(tableRowDataList2);
        list.add(tableData2);
        return list;
    }




    private static List<ChartData> getChartData() {
        List<ChartCategory> categoryDataList = new ArrayList<>();
        ChartCategory categoryData = new ChartCategory("第一季度", 8.2);
        ChartCategory categoryData2 = new ChartCategory("第二季度", 3.2);
        ChartCategory categoryData3 = new ChartCategory("第三季度", 2.6);
        categoryDataList.add(categoryData);
        categoryDataList.add(categoryData2);
        categoryDataList.add(categoryData3);

        List<ChartSeries> seriesDataList = new ArrayList<>();
        ChartSeries seriesData = new ChartSeries();
        seriesData.setSeriesName("销售额");
        seriesData.setChartCategoryList(categoryDataList);
        seriesDataList.add(seriesData);

        ChartData chartData = new ChartData();
        chartData.setChartSeriesList(seriesDataList);

        List<ChartData> chartDataList = new ArrayList<>();
        chartDataList.add(chartData);
        return chartDataList;
    }


    private static List<ChartData> getChartData2() {
        List<ChartCategory> categoryDataList = new ArrayList<>();
        ChartCategory categoryData = new ChartCategory("系列 1", 4.1);
        ChartCategory categoryData2 = new ChartCategory("系列 2", 3.0);
        ChartCategory categoryData3 = new ChartCategory("系列 3", 5.5);
        categoryDataList.add(categoryData);
        categoryDataList.add(categoryData2);
        categoryDataList.add(categoryData3);


        List<ChartCategory> categoryDataList1 = new ArrayList<>();
        ChartCategory categoryData1 = new ChartCategory("系列 1", 2.2);
        ChartCategory categoryData12 = new ChartCategory("系列 2", 6.3);
        ChartCategory categoryData13 = new ChartCategory("系列 3", 6.5);
        categoryDataList1.add(categoryData1);
        categoryDataList1.add(categoryData12);
        categoryDataList1.add(categoryData13);

        List<ChartCategory> categoryDataList2 = new ArrayList<>();
        ChartCategory categoryData21 = new ChartCategory("系列 1", 3.5);
        ChartCategory categoryData22 = new ChartCategory("系列 2", 4.7);
        ChartCategory categoryData23 = new ChartCategory("系列 3", 6.8);
        categoryDataList2.add(categoryData21);
        categoryDataList2.add(categoryData22);
        categoryDataList2.add(categoryData23);


        List<ChartCategory> categoryDataList3 = new ArrayList<>();
        ChartCategory categoryData31 = new ChartCategory("系列 1", 4.9);
        ChartCategory categoryData32 = new ChartCategory("系列 2", 3.4);
        ChartCategory categoryData33 = new ChartCategory("系列 3", 5.3);
        categoryDataList3.add(categoryData31);
        categoryDataList3.add(categoryData32);
        categoryDataList3.add(categoryData33);


        List<ChartSeries> seriesDataList = new ArrayList<>();
        ChartSeries seriesData = new ChartSeries();
        seriesData.setSeriesName("类别1");
        seriesData.setChartCategoryList(categoryDataList);

        ChartSeries seriesData1 = new ChartSeries();
        seriesData1.setSeriesName("类别2");
        seriesData1.setChartCategoryList(categoryDataList1);

        ChartSeries seriesData2 = new ChartSeries();
        seriesData2.setSeriesName("类别3");
        seriesData2.setChartCategoryList(categoryDataList2);

        ChartSeries seriesData3 = new ChartSeries();
        seriesData3.setSeriesName("类别4");
        seriesData3.setChartCategoryList(categoryDataList3);


        seriesDataList.add(seriesData);
        seriesDataList.add(seriesData1);
        seriesDataList.add(seriesData2);
        seriesDataList.add(seriesData3);


        ChartData chartData = new ChartData();
        chartData.setChartSeriesList(seriesDataList);

        List<ChartData> chartDataList = new ArrayList<>();
        chartDataList.add(chartData);
        return chartDataList;
    }


}
