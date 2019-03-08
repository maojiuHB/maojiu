package org.swkj.jhipster.web.rest;


import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author videon
 * @date 2018/12/17 下午3:03
 */
public class PIOUtilTest {
    private void build(File tmpFile, Map<String, String> contentMap, String exportFile) throws Exception {
        FileInputStream tempFileInputStream = new FileInputStream(tmpFile);
        HWPFDocument document = new HWPFDocument(tempFileInputStream);
        // 读取文本内容
        Range bodyRange = document.getRange();
        // 替换内容
        for (Map.Entry<String, String> entry : contentMap.entrySet()) {
            bodyRange.replaceText("${" + entry.getKey() + "}", entry.getValue());
        }
        //导出到文件
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        document.write(byteArrayOutputStream);
        OutputStream outputStream = new FileOutputStream(exportFile);
        outputStream.write(byteArrayOutputStream.toByteArray());
        outputStream.close();
    }
    @Test
    public void testExportWord2() throws Exception {
        String tmpFile = "/Users/videon/Desktop/template2.doc";
        String expFile = "/Users/videon/Desktop/test.doc";
        Map<String, String> datas = new HashMap<String, String>();
        datas.put("title1", "第二次会议");
        datas.put("title2", "第一次全体会议");
        datas.put("personNum", "50");
        datas.put("personTo", "30");
        datas.put("personNotTo", "20");
        datas.put("date","2018/12/12 18:30:00");


        datas.put("title","人大会议");
        datas.put("person","20");
        build(ResourceUtils.getFile(tmpFile), datas, expFile);
    }

    @Test
    public void docTest()   {
        String templatePath = "/Users/videon/Desktop/template1.docx";
        FileInputStream in = null;
        FileOutputStream out = null;
        try {
            in = new FileInputStream(templatePath);
            XWPFDocument doc = new XWPFDocument(in);
            //文本替换
            Map<String, String> param = new HashMap<String, String>();
            param.put("title1", "第二次会议");
            param.put("title2", "第一次全体会议");
            param.put("personNum", "50");
            param.put("personTo", "30");
            param.put("personNotTo", "20");
            param.put("date","2018/12/12 18:30:00");
            List<XWPFParagraph> allXWPFParagraphs = doc.getParagraphs();
            for (XWPFParagraph xwpfParagraph : allXWPFParagraphs) {
                List<XWPFRun> runs = xwpfParagraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (StringUtils.isNoneBlank(text)) {
                        if (text.equals("table")) {
                            //表格生成 6行5列.
                            int rows = 6;
                            int cols = 5;
                            XmlCursor cursor = xwpfParagraph.getCTP().newCursor();
                            XWPFTable tableOne = doc.insertNewTbl(cursor);

                            //样式控制
                            CTTbl ttbl = tableOne.getCTTbl();
                            CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
                            CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
                            CTJc cTJc = tblPr.addNewJc();
                            cTJc.setVal(STJc.Enum.forString("center"));//表格居中
                            tblWidth.setW(new BigInteger("9000"));//每个表格宽度
                            tblWidth.setType(STTblWidth.DXA);

                            //表格创建
                            XWPFTableRow tableRowTitle = tableOne.getRow(0);
                            tableRowTitle.getCell(0).setText("标题");
                            tableRowTitle.addNewTableCell().setText("内容");
                            tableRowTitle.addNewTableCell().setText("姓名");
                            tableRowTitle.addNewTableCell().setText("日期");
                            tableRowTitle.addNewTableCell().setText("备注");
                            for (int i = 1; i < rows; i++) {
                                XWPFTableRow createRow = tableOne.createRow();
                                for (int j = 0; j < cols; j++) {
                                    createRow.getCell(j).setText("我是第"+i+"行,第"+(j+1)+"列");
                                }
                            }
                            run.setText("", 0);
                        }else {
                            for (Map.Entry<String, String> entry : param.entrySet()) {
                                String key = entry.getKey();
                                if (text.indexOf(key) != -1) {
                                    text = text.replace(key, entry.getValue());
                                    run.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }

            out = new FileOutputStream("/Users/videon/Desktop/test.docx");
            // 输出
            doc.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (out != null) {
                    out.close();
                }
                if (in != null) {
                    in.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    
}
