import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.TextAlign;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

//        原文链接：https://blog.csdn.net/huangwenyi1010/article/details/51705402

/**
 * PPT简单插入表个
 */
public class PptTableTest {
    public static void main(String[] args) throws Exception{
        /** 文件路径 **/
        String templateFilePath = System.getProperty("user.dir") +
                File.separator + "src\\test\\diyun_land_report_template.pptx";

        /** 加载PPT **/
        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(templateFilePath));

        /** 创建一个slide，理解为PPT里的每一页 **/
//        XSLFSlide slide = ppt.createSlide();//先搜索是否有空白样式的ppt页面，默认用空白的样式创建一页PPT
        XSLFSlideLayout xslfSlideLayout = ppt.getSlideMasters().get(0).getSlideLayouts()[5];
        XSLFSlide slide = ppt.createSlide(xslfSlideLayout);//用指定页面的样式来创建一页PPT

        /** 获得slideMasters**/
        List<XSLFSlideMaster> slideMasters = ppt.getSlideMasters();
        /** 创建表格**/
        XSLFTable table = slide.createTable();
        /** 设置表格 x ,y ,width,height **/
        Rectangle2D rectangle2D = new Rectangle2D.Double(20,100,300,300);//X表示距离左侧的距离，y表示
        /** 生成第一行 **/
        XSLFTableRow firstRow =  table.addRow();
        /** 生成第一个单元格**/
        XSLFTableCell firstCell =  firstRow.addCell();

        /** 设置单元格的边框颜色 **/
        firstCell.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        firstCell.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        firstCell.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        firstCell.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));

        /** 设置单元格边框 **/
        firstCell.setBorderWidth(TableCell.BorderEdge.left, 3);
        firstCell.setBorderWidth(TableCell.BorderEdge.right, 3);
        firstCell.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        firstCell.setBorderWidth(TableCell.BorderEdge.top, 3);
        /** 设置文本 **/
        firstCell.setText("AAA");


        /** 设置单元格的边框宽度 **/
        XSLFTableCell secondCell =  firstRow.addCell();
        secondCell.setText("BBB");

        /** 设置单元格的边框颜色 **/
        secondCell.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        secondCell.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        secondCell.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        secondCell.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));

        /** 设置单元格边框 **/
        secondCell.setBorderWidth(TableCell.BorderEdge.left, 3);
        secondCell.setBorderWidth(TableCell.BorderEdge.right, 3);
        secondCell.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        secondCell.setBorderWidth(TableCell.BorderEdge.top, 3);

        /** 生成第一行 **/
        XSLFTableRow secondRow =  table.addRow();
        XSLFTableCell cell21 =  secondRow.addCell();
        cell21.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        cell21.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        cell21.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        cell21.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));
        cell21.setBorderWidth(TableCell.BorderEdge.left, 3);
        cell21.setBorderWidth(TableCell.BorderEdge.right, 3);
        cell21.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        cell21.setBorderWidth(TableCell.BorderEdge.top, 3);
        cell21.setText("CCC");

        XSLFTableCell cell22 =  secondRow.addCell();
        cell22.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        cell22.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        cell22.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        cell22.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));
        cell22.setBorderWidth(TableCell.BorderEdge.left, 3);
        cell22.setBorderWidth(TableCell.BorderEdge.right, 3);
        cell22.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        cell22.setBorderWidth(TableCell.BorderEdge.top, 3);
        cell22.setText("DDD");
        table.setAnchor(rectangle2D);





        /** 创建表格**/
        XSLFTable table2 = slide.createTable();
        /** 设置表格 x ,y ,width,height **/
        Rectangle2D rectangle2D2 = new Rectangle2D.Double(300,100,300,300);//X表示距离左侧的距离，y表示
        /** 生成第行 **/
        firstRow =  table2.addRow();

        /** 生成第一个单元格**/
        firstCell =  firstRow.addCell();


        /** 设置单元格的边框颜色 **/
        firstCell.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        firstCell.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        firstCell.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        firstCell.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));

        /** 设置单元格边框 **/
        firstCell.setBorderWidth(TableCell.BorderEdge.left, 3);
        firstCell.setBorderWidth(TableCell.BorderEdge.right, 3);
        firstCell.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        firstCell.setBorderWidth(TableCell.BorderEdge.top, 3);
        /** 设置文本 **/
        firstCell.setText("EEE");
//        firstCell.setVerticalAlignment(VerticalAlignment.MIDDLE);
//        firstCell.setHorizontalCentered(true);
//        firstCell.addNewTextParagraph().setTextAlign(TextParagraph.TextAlign.CENTER);




        /** 设置单元格的边框宽度 **/
        secondCell =  firstRow.addCell();

        /** 设置单元格的边框颜色 **/
        secondCell.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        secondCell.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        secondCell.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        secondCell.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));

        /** 设置单元格边框 **/
        secondCell.setBorderWidth(TableCell.BorderEdge.left, 3);
        secondCell.setBorderWidth(TableCell.BorderEdge.right, 3);
        secondCell.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        secondCell.setBorderWidth(TableCell.BorderEdge.top, 3);

        table2.mergeCells(0,0,0,1); //合并第一行的两个单元格



        /** 生成第二行 **/
        secondRow =  table2.addRow();
        cell21 =  secondRow.addCell();
        cell21.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        cell21.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        cell21.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        cell21.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));
        cell21.setBorderWidth(TableCell.BorderEdge.left, 3);
        cell21.setBorderWidth(TableCell.BorderEdge.right, 3);
        cell21.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        cell21.setBorderWidth(TableCell.BorderEdge.top, 3);
        cell21.setText("GGG");

        cell22 =  secondRow.addCell();
        cell22.setBorderColor(TableCell.BorderEdge.bottom,new Color(10,100,120));
        cell22.setBorderColor(TableCell.BorderEdge.right,new Color(10,100,120));
        cell22.setBorderColor(TableCell.BorderEdge.left,new Color(10,100,120));
        cell22.setBorderColor(TableCell.BorderEdge.top,new Color(10,100,120));
        cell22.setBorderWidth(TableCell.BorderEdge.left, 3);
        cell22.setBorderWidth(TableCell.BorderEdge.right, 3);
        cell22.setBorderWidth(TableCell.BorderEdge.bottom, 3);
        cell22.setBorderWidth(TableCell.BorderEdge.top, 3);
        cell22.setText("HHH");
        table2.setAnchor(rectangle2D2);


        /** 输出文件 **/
        ppt.write(new FileOutputStream(templateFilePath));
    }

}
