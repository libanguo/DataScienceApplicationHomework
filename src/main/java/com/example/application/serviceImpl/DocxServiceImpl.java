package com.example.application.serviceImpl;

import com.example.application.PO.*;
import com.example.application.service.DocxService;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

@Service
public class DocxServiceImpl implements DocxService {
    @Override
    public void wordParse(MultipartFile file, HashMap<String, List<ParagraphPO>> paragraphHashMap, HashMap<String, List<TablePO>> tableHashMap, HashMap<String, List<ImagePO>> imageHashMap, HashMap<String, List<TitlePO>> titleHashMap,HashMap<String,List<FontPO>> fontHashMap, String token) throws IOException {
        InputStream is =file.getInputStream();
        List<ParagraphPO> paragraphPOS=new ArrayList<>();
        List<TablePO> tablePOS=new ArrayList<>();
        List<ImagePO> imagePOS=new ArrayList<>();
        List<TitlePO> titlePOS=new ArrayList<>();
        List<FontPO> fontPOS=new ArrayList<>();
        XWPFDocument doc=new XWPFDocument(is);
        Iterator<IBodyElement> iterator=doc.getBodyElementsIterator();
        int paragraphId=1;
        Boolean tableFtlag=false;
        while (iterator.hasNext()){
            IBodyElement item=iterator.next();
            if(item instanceof XWPFParagraph){
                XWPFParagraph paragraph=(XWPFParagraph) item;
                ParagraphPO paragraphPO = new ParagraphPO();
                if(paragraph.getRuns().size()==0){
                    continue;
                }
                XWPFRun xwpfRun = paragraph.getRuns().get(0);
                paragraphPO.setParagraphId(paragraphId);
                if(tableFtlag){
                    TablePO tablePO=tablePOS.get(tablePOS.size()-1);
                    tablePO.setParagraphAfter(new TableGraphPO(paragraphId,paragraph.getParagraphText()));
                    if(paragraph.getParagraphText().length()<=10){
                        tablePO.setTextAfter(paragraph.getParagraphText());
                    }
                    tableFtlag=false;
                }
                paragraphPO.setParagraphText(paragraph.getParagraphText());
                paragraphPO.setFontSize(xwpfRun.getFontSize());
                paragraphPO.setFontName(xwpfRun.getFontName());
                paragraphPO.setIsBold(xwpfRun.isBold());
                paragraphPO.setIsItalic(xwpfRun.isItalic());
                FontPO fontPO=new FontPO();
                fontPO.setFontSize(xwpfRun.getFontSize());
                fontPO.setFontName(xwpfRun.getFontName());
                fontPO.setIsBold(xwpfRun.isBold());
                fontPO.setIsItalic(xwpfRun.isItalic());
                fontPO.setFontAlignment(paragraph.getFontAlignment());
                if(xwpfRun.getColor()!=null){
                    fontPO.setColor(Integer.parseInt(xwpfRun.getColor()));
                }
                fontPOS.add(fontPO);
                paragraphPO.setIsInTable(false);
                //TODO 存疑
                if(paragraph.getCTP().getPPr().getOutlineLvl()!=null){
                    paragraphPO.setLvl(Integer.parseInt(paragraph.getStyle()));
                }
                paragraphPO.setLineSpacing(paragraph.getSpacingLineRule().getValue());
                paragraphPO.setFontAlignment(paragraph.getFontAlignment());
                paragraphPO.setIsTableRowEnd(false);
                paragraphPO.setIndentFromLeft(paragraph.getIndentFromLeft());
                paragraphPO.setIndentFromRight(paragraph.getIndentFromRight());
                paragraphPOS.add(paragraphPO);
                if(paragraph.getCTP().getPPr().getOutlineLvl()!=null){
                    String style = paragraph.getStyle();
                    if (style.compareTo("9") < 0) {
                        TitlePO title = new TitlePO();
                        title.setParagraphText(paragraph.getParagraphText());
                        title.setParagraphId(paragraphId);
                        title.setIndentFromLeft(paragraph.getIndentFromLeft());
                        title.setIndentFromRight(paragraph.getIndentFromRight());
                        title.setFirstLineIndent(paragraph.getFirstLineIndent());
                        title.setLvl(Integer.parseInt(paragraph.getStyle()));
                        titlePOS.add(title);
                    }
                }
                for (XWPFRun xwpfRun1: paragraph.getRuns()){
                    for (XWPFPicture picture : xwpfRun1.getEmbeddedPictures()){
                        ImagePO imagePO=new ImagePO();
                        imagePO.setParagraphBefore(paragraphId-1);
                        imagePO.setWidth(picture.getWidth());
                        imagePO.setHeight(picture.getDepth());
                        imagePO.setTextBefore(picture.getDescription());
                        imagePO.setTextAfter(picture.getDescription());
                        imagePO.setSuggestFileExtension(picture.getPictureData().suggestFileExtension());
                        imagePO.setFilename(picture.getPictureData().getFileName());
                        imagePO.setBase64Content(picture.getPictureData().getData());
                        imagePOS.add(imagePO);
                    }
                }
                paragraphId++;
            }
            else if(item instanceof XWPFTable){
                TablePO tablePO=new TablePO();
                List<TableGraphPO> tableGraphPOS=new ArrayList<>();
                XWPFTable table=(XWPFTable) item;
                List<XWPFTableRow> rows = table.getRows();
                if(paragraphId>1){
                    ParagraphPO tmp=paragraphPOS.get(paragraphPOS.size()-1);
                    tablePO.setParagraphBefore(new TableGraphPO(paragraphId-1,tmp.getParagraphText()));
                    if(tmp.getParagraphText().length()<=10){
                        tablePO.setTextBefore(tmp.getParagraphText());
                    }
                }
                for (XWPFTableRow row : rows) {
                    // 获取表格的每个单元格
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    for (XWPFTableCell cell : tableCells) {
                        // 获取单元格的内容
                        String text = cell.getText();
                        XWPFParagraph paragraph=cell.getParagraphs().get(0);
                        ParagraphPO paragraphPO=new ParagraphPO();
                        XWPFRun xwpfRun = paragraph.getRuns().get(0);
                        paragraphPO.setParagraphId(paragraphId);
                        paragraphPO.setParagraphText(text);
                        paragraphPO.setFontSize(xwpfRun.getFontSize());
                        paragraphPO.setFontName(xwpfRun.getFontName());
                        paragraphPO.setIsBold(xwpfRun.isBold());
                        paragraphPO.setIsItalic(xwpfRun.isItalic());
                        paragraphPO.setIsInTable(false);
                        //TODO 存疑
                        if(paragraph.getCTP().getPPr().getOutlineLvl()!=null){
                            paragraphPO.setLvl(Integer.parseInt(paragraph.getStyle()));
                        }
                        paragraphPO.setLineSpacing(paragraph.getSpacingLineRule().getValue());
                        paragraphPO.setFontAlignment(paragraph.getFontAlignment());
                        paragraphPO.setIsTableRowEnd(false);
                        paragraphPO.setIndentFromLeft(paragraph.getIndentFromLeft());
                        paragraphPO.setIndentFromRight(paragraph.getIndentFromRight());
                        paragraphPOS.add(paragraphPO);
                        TableGraphPO tableGraphPO=new TableGraphPO(paragraphId,text);
                        tableGraphPOS.add(tableGraphPO);
                        FontPO fontPO=new FontPO();
                        fontPO.setFontSize(xwpfRun.getFontSize());
                        fontPO.setFontName(xwpfRun.getFontName());
                        fontPO.setIsBold(xwpfRun.isBold());
                        fontPO.setIsItalic(xwpfRun.isItalic());
                        fontPO.setFontAlignment(paragraph.getFontAlignment());
                        if(xwpfRun.getColor()!=null){
                            fontPO.setColor(Integer.parseInt(xwpfRun.getColor()));
                        }
                        fontPOS.add(fontPO);
                        paragraphId++;
                    }
                }
                tablePO.setTableContent(tableGraphPOS);
                tableFtlag=true;
                tablePOS.add(tablePO);
            }
            else if(item instanceof XWPFPicture){
                XWPFPicture picture=(XWPFPicture) item;
                ImagePO imagePO=new ImagePO();
                imagePO.setParagraphBefore(paragraphId-1);
                imagePO.setWidth(picture.getWidth());
                imagePO.setHeight(picture.getDepth());
                imagePO.setTextBefore(picture.getDescription());
                imagePO.setTextAfter(picture.getDescription());
                imagePO.setSuggestFileExtension(picture.getPictureData().suggestFileExtension());
                imagePO.setFilename(picture.getPictureData().getFileName());
                imagePO.setBase64Content(picture.getPictureData().getData());
                imagePOS.add(imagePO);
            }
        }
        paragraphHashMap.put(token,paragraphPOS);
        tableHashMap.put(token,tablePOS);
        imageHashMap.put(token,imagePOS);
        titleHashMap.put(token,titlePOS);
        fontHashMap.put(token,fontPOS);
    }


    @Override
    public ParagraphFormatPO getParagraphFormat(ParagraphPO paragraph) {
        ParagraphFormatPO paragraphFormatPO = new ParagraphFormatPO();
        paragraphFormatPO.setLvl(paragraph.getLvl());
        paragraphFormatPO.setIndentFromLeft(paragraph.getIndentFromLeft());
        paragraphFormatPO.setIndentFromRight(paragraph.getIndentFromRight());
        paragraphFormatPO.setFirstLineIndent(paragraph.getFirstLineIndent());
        paragraphFormatPO.setLineSpacing(paragraph.getLineSpacing());
        return paragraphFormatPO;
    }

}
