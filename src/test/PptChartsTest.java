import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;


public class PptChartsTest {
    public static void main(String[] args) throws Exception {

        String template = "e:\\pie-chart-template.pptx";
        String path = "e:\\pie-chart-out.pptx";

        XMLSlideShow pptx = null;
        try {
            String chartTitle ="chart title";
            String[] names = new String[] {"East", "Middle", "West"};
            String[] values = new String[] {"189", "412", "250"};
            //打开模板ppt
            pptx = new XMLSlideShow(new FileInputStream(template));
            //获取第一个ppt页面
            XSLFSlide slide = pptx.getSlides().get(0);
            //遍历第一页元素找到图表
            XSLFChart chart = null;
            for(POIXMLDocumentPart part : slide.getRelations()){
                if(part instanceof XSLFChart){
                    chart = (XSLFChart) part;
                    break;
                }
            }
            if (chart == null) {
                System.out.println("no chart");
                return ;
            }

            POIXMLDocumentPart xlsPart = chart.getRelations().get(0);

            //把图表绑定到Excel workbook中
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet = wb.createSheet();

            CTChart ctChart = chart.getCTChart();
            CTPlotArea plotArea = ctChart.getPlotArea();

            CTPieChart pieChart = plotArea.getPieChartArray(0);
            //获取图表的系列
            CTPieSer ser = pieChart.getSerArray(0);

            // Series Text
            CTSerTx tx = ser.getTx();
            tx.getStrRef().getStrCache().getPtArray(0).setV(chartTitle);
            sheet.createRow(0).createCell(1).setCellValue(chartTitle);
            String titleRef = new CellReference(sheet.getSheetName(), 0, 1, true, true).formatAsString();
            tx.getStrRef().setF(titleRef);

            // Category Axis Data
            CTAxDataSource cat = ser.getCat();
            CTStrData strData = cat.getStrRef().getStrCache();

            //获取图表的值
            CTNumDataSource val = ser.getVal();
            CTNumData numData = val.getNumRef().getNumCache();

            strData.setPtArray(null);  // unset old axis text
            numData.setPtArray(null);  // unset old values

            // set model
            int idx = 0;
            int rownum = 1;
            String ln;
            for (int i=0; i<names.length; i++) {
                CTNumVal numVal = numData.addNewPt();
                numVal.setIdx(idx);
                numVal.setV(values[i]);

                CTStrVal sVal = strData.addNewPt();
                sVal.setIdx(idx);
                sVal.setV(names[i]);

                idx++;
                XSSFRow row = sheet.createRow(rownum++);
                row.createCell(0).setCellValue(names[i]);
                row.createCell(1).setCellValue(Double.valueOf(values[i]));
            }
            numData.getPtCount().setVal(idx);
            strData.getPtCount().setVal(idx);

            String numDataRange = new CellRangeAddress(1, rownum-1, 1, 1).formatAsString(sheet.getSheetName(), true);
            val.getNumRef().setF(numDataRange);
            String axisDataRange = new CellRangeAddress(1, rownum-1, 0, 0).formatAsString(sheet.getSheetName(), true);
            cat.getStrRef().setF(axisDataRange);

            //更新嵌入的workbook
            OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();
            wb.write(xlsOut);
            xlsOut.close();

            //保存文件
            OutputStream out = new FileOutputStream(path);
            pptx.write(out);
            out.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (pptx != null) {
                try {
                    pptx.close();
                } catch (Exception ee){
                    //nothing
                }
            }
        }

//        原文链接：https://blog.csdn.net/starandsea/article/details/51741328

    }
}
