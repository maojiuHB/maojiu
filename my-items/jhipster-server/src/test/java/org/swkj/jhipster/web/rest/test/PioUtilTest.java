package org.swkj.jhipster.web.rest.test;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author videon
 * @date 2018/12/17 下午8:50
 */
public class PioUtilTest {
    @Test
    public void export(){
        List<String> list = new ArrayList<>();
        list.add("第1次全体次会议");
        list.add("第2次会议");
        list.add("缺席名单");
        list.add("应到 50 人");
        list.add("实到 10 人");
        list.add("未到 40 人");
        list.add("");
        list.add(new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format(new Date()));
        try {
            exportWord(list
            );
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlException e) {
            e.printStackTrace();
        }

    }
    private void  exportWord(List<String> list) throws IOException, XmlException {
        XWPFDocument document = new XWPFDocument();
        // 输出word的位置
        FileOutputStream out = new FileOutputStream(new File("/Users/videon/Desktop/1.docx"));
        createParagraph(document,list);
        createTable(document,list);
        document.write(out);
        out.close();
    }
    private void createParagraph(XWPFDocument document, List<String> list){
        list.forEach(data ->{
            XWPFParagraph paragraph = document.createParagraph();
            // 段落居中
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText(data.toString());
            run.setFontSize(24);
            // 换行
            XWPFParagraph paragraph1 = document.createParagraph();
            XWPFRun paragraphRun1 = paragraph1.createRun();
            paragraphRun1.setText("\r");
        });
    }
    private void createTable(XWPFDocument document, List<String> list){
        //基本信息表格
        XWPFTable infoTable = document.createTable();
        //去表格边框
        infoTable.getCTTbl().getTblPr().unsetTblBorders();
        //列宽自动分割
        CTTblWidth infoTableWidth = infoTable.getCTTbl().addNewTblPr().addNewTblW();
        infoTableWidth.setType(STTblWidth.DXA);
        infoTableWidth.setW(BigInteger.valueOf(9072));
        int a = list.size() % 3;
        if (a == 1){
            list.add("");
            list.add("");
        }else if (a == 2){
            list.add("");
        }
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
    }
}
