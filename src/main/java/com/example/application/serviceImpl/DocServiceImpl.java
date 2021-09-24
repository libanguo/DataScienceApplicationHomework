package com.example.application.serviceImpl;


import com.example.application.PO.*;
import com.example.application.service.DocService;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;


@Service
public class DocServiceImpl implements DocService {

    @Override
    public List<ParagraphPO> getAllParagraphs(MultipartFile file) throws IOException {
        List<ParagraphPO> paragraphList = new ArrayList<>();
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        int paraNum = range.numParagraphs();
        for (int i = 0; i < paraNum; i++) {
            ParagraphPO paragraphPO = new ParagraphPO();
            Paragraph paragraph = range.getParagraph(i);
            CharacterRun characterRun = paragraph.getCharacterRun(0);
            paragraphPO.setParagraphText("" + paragraph.getIlfo() + " " + paragraph.text());
            int id = i + 1;
            paragraphPO.setParagraphId(paragraph.getIlfo() == 0 ? id : paragraph.getIlfo());
            paragraphPO.setFontName(characterRun.getFontName());
            paragraphPO.setFontSize(characterRun.getFontSize());
            paragraphPO.setFontAlignment(paragraph.getFontAlignment());
            paragraphPO.setFirstLineIndent(paragraph.getFirstLineIndent());
            paragraphPO.setIsBold(characterRun.isBold());
            paragraphPO.setIndentFromLeft(paragraph.getIndentFromLeft());
            paragraphPO.setIndentFromRight(paragraph.getIndentFromRight());
            paragraphPO.setIsItalic(characterRun.isItalic());
            paragraphPO.setLineSpacing((65536 - paragraph.getLineSpacing().toInt()));
            paragraphPO.setIsInTable(paragraph.isInTable());
            paragraphPO.setLvl(paragraph.getLvl());
            paragraphPO.setIsTableRowEnd(paragraph.isTableRowEnd());
            paragraphList.add(paragraphPO);
        }
        return paragraphList;
    }

    @Override
    public List<TablePO> getAllTables(MultipartFile file) throws IOException {
        List<TablePO> tableList = new ArrayList<>();
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        int paraNum = range.numParagraphs();
        // 从头开始遍历所有的段，如果在table里，co为i，然后用co去找表格段尾
        int co = 0;
        int i=0;
        while (i<paraNum){
            Paragraph paragraph = range.getParagraph(i);
            TablePO tablePO = new TablePO();
            if (paragraph.isInTable()) {
                co = i;
                while (co<paraNum && range.getParagraph(co).isInTable()) {
                    co++;
                }
                if (i != 0 && co != paraNum) {
                    Paragraph paragraphPref = range.getParagraph(i - 1);
                    Paragraph paragraphAfter = range.getParagraph(co);
                    TableGraphPO paragraphPOPref = new TableGraphPO();
                    TableGraphPO paragraphPOAfter = new TableGraphPO();
                    paragraphPOPref.setParagraphId(i - 1);
                    paragraphPOPref.setTableTextContent(paragraphPref.text());
                    paragraphPOAfter.setParagraphId(co);
                    paragraphPOAfter.setTableTextContent(paragraphAfter.text());
                    tablePO.setParagraphBefore(paragraphPOPref);
                    tablePO.setParagraphAfter(paragraphPOAfter);
                    if (paragraphPOPref.getTableTextContent().length() <= 10) {
                        tablePO.setTextBefore(paragraphPOPref.getTableTextContent());
                    } else {
                        tablePO.setTextBefore("");
                    }
                    if (paragraphPOAfter.getTableTextContent().length() <= 10) {
                        tablePO.setTextAfter(paragraphPOAfter.getTableTextContent());
                    } else {
                        tablePO.setTextAfter("");
                    }
                } else if (i == 0 && co != paraNum) {
                    TableGraphPO paragraphPOPref = new TableGraphPO();
                    TableGraphPO paragraphPOAfter = new TableGraphPO();
                    Paragraph paragraphAfter = range.getParagraph(co);
                    paragraphPOPref.setParagraphId(i - 1);
                    paragraphPOPref.setTableTextContent("");
                    paragraphPOAfter.setParagraphId(co);
                    paragraphPOAfter.setTableTextContent(paragraphAfter.text());
                    tablePO.setParagraphBefore(paragraphPOPref);
                    tablePO.setParagraphAfter(paragraphPOAfter);
                    if (paragraphPOPref.getTableTextContent().length() <= 10) {
                        tablePO.setTextBefore(paragraphPOPref.getTableTextContent());
                    } else {
                        tablePO.setTextBefore("");
                    }
                    if (paragraphPOAfter.getTableTextContent().length() <= 10) {
                        tablePO.setTextAfter(paragraphPOAfter.getTableTextContent());
                    } else {
                        tablePO.setTextAfter("");
                    }
                } else if (co == paraNum && i != 0) {
                    Paragraph paragraphPref = range.getParagraph(i - 1);
                    TableGraphPO paragraphPOPref = new TableGraphPO();
                    TableGraphPO paragraphPOAfter = new TableGraphPO();
                    paragraphPOPref.setParagraphId(i - 1);
                    paragraphPOPref.setTableTextContent(paragraphPref.text());
                    paragraphPOAfter.setParagraphId(co);
                    paragraphPOAfter.setTableTextContent("");
                    tablePO.setParagraphBefore(paragraphPOPref);
                    tablePO.setParagraphAfter(paragraphPOAfter);
                    if (paragraphPOPref.getTableTextContent().length() <= 10) {
                        tablePO.setTextBefore(paragraphPOPref.getTableTextContent());
                    } else {
                        tablePO.setTextBefore("");
                    }
                    if (paragraphPOAfter.getTableTextContent().length() <= 10) {
                        tablePO.setTextAfter(paragraphPOAfter.getTableTextContent());
                    } else {
                        tablePO.setTextAfter("");
                    }
                }
                List<TableGraphPO> tableContent = new ArrayList<>();
                for (int j = i; j < co; j++) {
                    TableGraphPO tablePOTemp = new TableGraphPO();
                    Paragraph paragraphTemp = range.getParagraph(j);
                    tablePOTemp.setParagraphId(j);
                    tablePOTemp.setTableTextContent(paragraphTemp.text());
                    tableContent.add(tablePOTemp);
                }
                tablePO.setTableContent(tableContent);
                tableList.add(tablePO);
                i=co;
            }
            else {
                i++;
            }
        }
//        for (int i = co; i < paraNum; i++) {
//
//        }
        return tableList;
    }

    @Override
    public List<ImagePO> getAllImages(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        int length = doc.characterLength();
        PicturesTable pTable = doc.getPicturesTable();
        List<ImagePO> imagePOList = new ArrayList<>();
        for (int i = 0; i < length; i++) {
            Range range = new Range(i, i + 1, doc);
            CharacterRun cr = range.getCharacterRun(0);
            if (pTable.hasPicture(cr)) {
                ImagePO imagePO = new ImagePO();
                Picture pic = pTable.extractPicture(cr, false);
                imagePO.setFilename(pic.suggestFullFileName());
                imagePO.setTextBefore(pic.getDescription());
                imagePO.setTextAfter(pic.getDescription());
                imagePO.setBase64Content(pic.getContent());
                imagePO.setHeight(Double.parseDouble(pic.getHeight()+""));
                imagePO.setWidth(Double.parseDouble(pic.getWidth()+""));
                imagePO.setSuggestFileExtension(pic.suggestFileExtension());
                imagePOList.add(imagePO);
            }
        }
        return imagePOList;
    }

    @Override
    public List<TitlePO> getAllTitles(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        int paraNum = range.numParagraphs();
        List<TitlePO> titlePOS = new ArrayList<>();
        for (int i = 0; i < paraNum; i++) {
            Paragraph paragraph = range.getParagraph(i);
            if (paragraph.getLvl() < 9) {
                TitlePO titlePO = new TitlePO();
                titlePO.setLvl(paragraph.getLvl());
                titlePO.setParagraphText(paragraph.text());
                titlePO.setLineSpacing((65536 - paragraph.getLineSpacing().toInt()) / 20);
                titlePO.setIndentFromLeft(paragraph.getIndentFromLeft());
                titlePO.setIndentFromRight(paragraph.getIndentFromRight());
                titlePO.setFirstLineIndent(paragraph.getFirstLineIndent());
                titlePOS.add(titlePO);
            }
        }
        return titlePOS;
    }

    @Override
    public ParagraphPO getParagraphText(Paragraph paragraph, int id) {
        Paragraph paragraphToGet = paragraph;
        ParagraphPO paragraphPO = new ParagraphPO();
        CharacterRun characterRun = paragraphToGet.getCharacterRun(0);
        paragraphPO.setParagraphText("" + paragraphToGet.getIlfo() + " " + paragraphToGet.text());
        paragraphPO.setParagraphId(paragraphToGet.getIlfo() == 0 ? id : paragraphToGet.getIlfo());
        paragraphPO.setFontName(characterRun.getFontName());
        paragraphPO.setFontSize(characterRun.getFontSize());
        paragraphPO.setFontAlignment(paragraphToGet.getFontAlignment());
        paragraphPO.setFirstLineIndent(paragraphToGet.getFirstLineIndent());
        paragraphPO.setIsBold(characterRun.isBold());
        paragraphPO.setIndentFromLeft(paragraphToGet.getIndentFromLeft());
        paragraphPO.setIndentFromRight(paragraphToGet.getIndentFromRight());
        paragraphPO.setIsItalic(characterRun.isItalic());
        paragraphPO.setLineSpacing((65536 - paragraphToGet.getLineSpacing().toInt()) / 20);
        paragraphPO.setIsInTable(paragraphToGet.isInTable());
        paragraphPO.setLvl(paragraphToGet.getLvl());
        paragraphPO.setIsTableRowEnd(paragraphToGet.isTableRowEnd());
        return paragraphPO;
    }

    @Override
    public ParagraphFormatPO getParagraphFormat(Paragraph paragraph, int id) {
        Paragraph paragraphToGet = paragraph;
        ParagraphFormatPO paragraphFormatPO = new ParagraphFormatPO();
        paragraphFormatPO.setLineSpacing(65536 - paragraphToGet.getLineSpacing().toInt());
        paragraphFormatPO.setIndentFromLeft(paragraphToGet.getIndentFromLeft());
        paragraphFormatPO.setIndentFromRight(paragraphToGet.getIndentFromRight());
        paragraphFormatPO.setFirstLineIndent(paragraphToGet.getFirstLineIndent());
        paragraphFormatPO.setLvl(paragraphToGet.getLvl());
        return paragraphFormatPO;
    }

    @Override
    public FontPO getParagraphFontFormat(Paragraph paragraph) {
        Paragraph paragraphToGet = paragraph;
        FontPO fontPO = new FontPO();
        CharacterRun characterRun = paragraphToGet.getCharacterRun(0);
        fontPO.setColor(characterRun.getColor()+"");
        fontPO.setFontSize(characterRun.getFontSize());
        fontPO.setFontName(characterRun.getFontName());
        fontPO.setFontAlignment(paragraphToGet.getFontAlignment());
        fontPO.setIsBold(characterRun.isBold());
        fontPO.setIsItalic(characterRun.isItalic());
        return fontPO;
    }

    @Override
    public List<ParagraphPO> getParagraphByTitle(MultipartFile file, int paragraphId) throws IOException {
        paragraphId--;
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        int paraNum = range.numParagraphs();
        int co = paragraphId;
        for (int i = paragraphId+1; i < paraNum; i++) {
            //1-8 9是段落，越大等级越小
            if (range.getParagraph(i).getLvl() > range.getParagraph(paragraphId).getLvl()) {
                co++;
            }
            else {
                break;
            }
        }
        List<ParagraphPO> paragraphList = new ArrayList<>();
        for (int i = paragraphId; i < co+1; i++) {
            ParagraphPO paragraphPO = new ParagraphPO();
            Paragraph paragraph = range.getParagraph(i);
            CharacterRun characterRun = paragraph.getCharacterRun(0);
            paragraphPO.setParagraphText("" + paragraph.getIlfo() + " " + paragraph.text());
            int id = i + 1;
            paragraphPO.setParagraphId(paragraph.getIlfo() == 0 ? id : paragraph.getIlfo());
            paragraphPO.setFontName(characterRun.getFontName());
            paragraphPO.setFontSize(characterRun.getFontSize());
            paragraphPO.setFontAlignment(paragraph.getFontAlignment());
            paragraphPO.setFirstLineIndent(paragraph.getFirstLineIndent());
            paragraphPO.setIsBold(characterRun.isBold());
            paragraphPO.setIndentFromLeft(paragraph.getIndentFromLeft());
            paragraphPO.setIndentFromRight(paragraph.getIndentFromRight());
            paragraphPO.setIsItalic(characterRun.isItalic());
            paragraphPO.setLineSpacing((65536 - paragraph.getLineSpacing().toInt()));
            paragraphPO.setIsInTable(paragraph.isInTable());
            paragraphPO.setLvl(paragraph.getLvl());
            paragraphPO.setIsTableRowEnd(paragraph.isTableRowEnd());
            paragraphList.add(paragraphPO);
        }
        return paragraphList;
    }

    @Override
    public List<ImagePO> getImagesByTitle(MultipartFile file, int paragraphId) throws IOException {
        paragraphId--;
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        int paraNum = range.numParagraphs();
        int co = paragraphId;
        for (int i = paragraphId+1; i < paraNum; i++) {
            //1-8 9是段落，越大等级越小
            if (range.getParagraph(i).getLvl() > range.getParagraph(paragraphId).getLvl()) {
                co++;
            }
            else {
                break;
            }
        }
        range = new Range(paragraphId, co, doc);
        int length = range.numParagraphs();
        List<ImagePO> imagePOList = new ArrayList<>();
        PicturesTable pTable = doc.getPicturesTable();
        for (int i = paragraphId; i < paragraphId + length+1; i++) {
            Range rangeTemp = new Range(i, i + 1, doc);
            CharacterRun cr = rangeTemp.getCharacterRun(0);
            if (pTable.hasPicture(cr)) {
                ImagePO imagePO = new ImagePO();
                Picture pic = pTable.extractPicture(cr, false);
                imagePO.setFilename(pic.suggestFullFileName());
                imagePO.setTextBefore(pic.getDescription());
                imagePO.setTextAfter(pic.getDescription());
                imagePO.setBase64Content(pic.getContent());
                imagePO.setHeight(Double.parseDouble(pic.getHeight()+""));
                imagePO.setWidth(Double.parseDouble(pic.getWidth()+""));
                imagePO.setSuggestFileExtension(pic.suggestFileExtension());
                imagePOList.add(imagePO);
            }
        }
        return imagePOList;
    }

    @Override
    public List<TablePO> getTablesByTitle(MultipartFile file, int paragraphId) throws IOException {
        paragraphId--;
        List<TablePO> tableList = new ArrayList<>();
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        int paraNumTemp = range.numParagraphs();
        int temp = paragraphId;
        for (int i = paragraphId+1; i < paraNumTemp; i++) {
            //1-8 9是段落，越大等级越小
            if (range.getParagraph(i).getLvl() > range.getParagraph(paragraphId).getLvl()) {
                temp++;
            }
            else {
                break;
            }
        }
        // 从标题开始遍历所有的段，如果在table里，co为i，然后用co去找表格段尾
        int co = paragraphId;
        for (int i = co; i < temp+1; i++) {
            Paragraph paragraph = range.getParagraph(i);
            TablePO tablePO = new TablePO();
            if (paragraph.isInTable()) {
                co = i;
                while (co<=temp && range.getParagraph(co).isInTable()) {
                    co++;
                }
                if (i != 0 && co != temp) {
                    Paragraph paragraphPref = range.getParagraph(i - 1);
                    Paragraph paragraphAfter = range.getParagraph(co);
                    TableGraphPO paragraphPOPref = new TableGraphPO();
                    TableGraphPO paragraphPOAfter = new TableGraphPO();
                    paragraphPOPref.setParagraphId(i - 1);
                    paragraphPOPref.setTableTextContent(paragraphPref.text());
                    paragraphPOAfter.setParagraphId(co);
                    paragraphPOAfter.setTableTextContent(paragraphAfter.text());
                    tablePO.setParagraphBefore(paragraphPOPref);
                    tablePO.setParagraphAfter(paragraphPOAfter);
                    if (paragraphPOPref.getTableTextContent().length() <= 10) {
                        tablePO.setTextBefore(paragraphPOPref.getTableTextContent());
                    } else {
                        tablePO.setTextBefore("");
                    }
                    if (paragraphPOAfter.getTableTextContent().length() <= 10) {
                        tablePO.setTextAfter(paragraphPOAfter.getTableTextContent());
                    } else {
                        tablePO.setTextAfter("");
                    }
                } else if (i == 0 && co != temp) {
                    TableGraphPO paragraphPOPref = new TableGraphPO();
                    TableGraphPO paragraphPOAfter = new TableGraphPO();
                    Paragraph paragraphAfter = range.getParagraph(co);
                    paragraphPOPref.setParagraphId(i - 1);
                    paragraphPOPref.setTableTextContent("");
                    paragraphPOAfter.setParagraphId(co);
                    paragraphPOAfter.setTableTextContent(paragraphAfter.text());
                    tablePO.setParagraphBefore(paragraphPOPref);
                    tablePO.setParagraphAfter(paragraphPOAfter);
                    if (paragraphPOPref.getTableTextContent().length() <= 10) {
                        tablePO.setTextBefore(paragraphPOPref.getTableTextContent());
                    } else {
                        tablePO.setTextBefore("");
                    }
                    if (paragraphPOAfter.getTableTextContent().length() <= 10) {
                        tablePO.setTextAfter(paragraphPOAfter.getTableTextContent());
                    } else {
                        tablePO.setTextAfter("");
                    }
                } else if (co == temp && i != 0) {
                    Paragraph paragraphPref = range.getParagraph(i - 1);
                    TableGraphPO paragraphPOPref = new TableGraphPO();
                    TableGraphPO paragraphPOAfter = new TableGraphPO();
                    paragraphPOPref.setParagraphId(i - 1);
                    paragraphPOPref.setTableTextContent(paragraphPref.text());
                    paragraphPOAfter.setParagraphId(co);
                    paragraphPOAfter.setTableTextContent("");
                    tablePO.setParagraphBefore(paragraphPOPref);
                    tablePO.setParagraphAfter(paragraphPOAfter);
                    if (paragraphPOPref.getTableTextContent().length() <= 10) {
                        tablePO.setTextBefore(paragraphPOPref.getTableTextContent());
                    } else {
                        tablePO.setTextBefore("");
                    }
                    if (paragraphPOAfter.getTableTextContent().length() <= 10) {
                        tablePO.setTextAfter(paragraphPOAfter.getTableTextContent());
                    } else {
                        tablePO.setTextAfter("");
                    }
                }
                List<TableGraphPO> tableContent = new ArrayList<>();
                for (int j = i; j < co; j++) {
                    TableGraphPO tablePOTemp = new TableGraphPO();
                    Paragraph paragraphTemp = range.getParagraph(j);
                    tablePOTemp.setParagraphId(j);
                    tablePOTemp.setTableTextContent(paragraphTemp.text());
                    tableContent.add(tablePOTemp);
                }
                tablePO.setTableContent(tableContent);
                tableList.add(tablePO);
                i=co-1;
            }
        }
        return tableList;
    }


}
