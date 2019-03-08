package org.swkj.jhipster.web.rest.test;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author videon
 * @date 2018/12/17 下午7:30
 */
public class PioTest {
    @Test
    public void write2Docx()throws Exception{
        XWPFDocument document= new XWPFDocument();


        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("/Users/videon/Desktop/test.docx"));
        List list = new ArrayList();
        list.add("第2次会议");
        list.add("第1次全体次会议");
        list.add("缺席名单");
        list.add("应到 50 人");
        list.add("实到 10 人");
        list.add("未到 40 人");
        list.add("2018/12/11 19:20:22");


        list.forEach(data ->{
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText(data.toString());
            run.setFontSize(20);
            XWPFParagraph paragraph1 = document.createParagraph();
            XWPFRun paragraphRun1 = paragraph1.createRun();
            paragraphRun1.setText("\r");
        });

        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();
        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun titleParagraphRun = titleParagraph.createRun();



        //设置段落背景颜色
        CTShd cTShd = titleParagraphRun.getCTR().addNewRPr().addNewShd();
        cTShd.setVal(STShd.CLEAR);
        cTShd.setFill("97FFFF");

        //换行
        XWPFParagraph paragraph1 = document.createParagraph();
        XWPFRun paragraphRun1 = paragraph1.createRun();
        paragraphRun1.setText("\r");


        //基本信息表格
        XWPFTable infoTable = document.createTable();
        //去表格边框
        infoTable.getCTTbl().getTblPr().unsetTblBorders();

        //列宽自动分割
        CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
        infoTableWidth.setType(STTblWidth.DXA);
        infoTableWidth.setW(BigInteger.valueOf(9072));
        int a = list.size() % 3;
        System.out.println("1"+list.size());
        if (a == 1){
            list.add("");
            list.add("");
        }else if (a == 2){
            list.add("");
        }
        System.out.println("2"+list.size());
        for (int i = 0; i < list.size(); i++){
            if (i < 3){
                XWPFTableRow infoTableRowOne = infoTable.getRow(0);
                infoTableRowOne.getCell(0).setText(list.get(i).toString());
                infoTableRowOne.addNewTableCell().setText(list.get(++i).toString());
                infoTableRowOne.addNewTableCell().setText(list.get(++i).toString());
            }else {
                XWPFTableRow infoTableRow = infoTable.createRow();
                infoTableRow.getCell(0).setText(list.get(i).toString());
                infoTableRow.getCell(1).setText(list.get(++i).toString());
                infoTableRow.getCell(2).setText(list.get(++i).toString());
            }
        }

//        //表格第一行
//        XWPFTableRow infoTableRowOne = infoTable.getRow(0);
//        infoTableRowOne.getCell(0).setText("职位");
//        infoTableRowOne.addNewTableCell().setText(": Java 开发工程师");
//        infoTableRowOne.addNewTableCell().setText("jkdlsaj");
//
//        //表格第二行
//        XWPFTableRow infoTableRowTwo = infoTable.createRow();
//        infoTableRowTwo.getCell(0).setText("姓名");
//        infoTableRowTwo.getCell(1).setText(": seawater");
//
//        //表格第三行
//        XWPFTableRow infoTableRowThree = infoTable.createRow();
//        infoTableRowThree.getCell(0).setText("生日");
//        infoTableRowThree.getCell(1).setText(": xxx-xx-xx");
//
//        //表格第四行
//        XWPFTableRow infoTableRowFour = infoTable.createRow();
//        infoTableRowFour.getCell(0).setText("性别");
//        infoTableRowFour.getCell(1).setText(": 男");
//
//        //表格第五行
//        XWPFTableRow infoTableRowFive = infoTable.createRow();
//        infoTableRowFive.getCell(0).setText("现居地");
//        infoTableRowFive.getCell(1).setText(": xx");
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);

        //添加页眉
        CTP ctpHeader = CTP.Factory.newInstance();
        CTR ctrHeader = ctpHeader.addNewR();
        CTText ctHeader = ctrHeader.addNewT();
        String headerText = "ctpHeader";
        ctHeader.setStringValue(headerText);
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeader, document);
        //设置为右对齐
        headerParagraph.setAlignment(ParagraphAlignment.RIGHT);
        XWPFParagraph[] parsHeader = new XWPFParagraph[1];
        parsHeader[0] = headerParagraph;
        policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, parsHeader);

        //添加页脚
        CTP ctpFooter = CTP.Factory.newInstance();
        CTR ctrFooter = ctpFooter.addNewR();
        CTText ctFooter = ctrFooter.addNewT();
        String footerText = "ctpFooter";
        ctFooter.setStringValue(footerText);
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooter, document);
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFParagraph[] parsFooter = new XWPFParagraph[1];
        parsFooter[0] = footerParagraph;
        policy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, parsFooter);

        document.write(out);
        out.close();
    }
}
