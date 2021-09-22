package com.example.application.serviceImpl;


import com.example.application.PO.*;
import com.example.application.service.DocService;
import org.apache.poi.hwpf.HWPFDocument;
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
            //TODO
            paragraphPO.setLineSpacing((65536-paragraph.getLineSpacing().toInt())/20);
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
        //获取内容和编号
        int co = 0;
        for (int i = co; i < paraNum; i++) {
            Paragraph paragraph = range.getParagraph(i);
            TablePO tablePO = new TablePO();
            if (paragraph.isInTable()) {
                co = i;
                while (range.getParagraph(co).isInTable()) {
                    co++;
                }
                if (i != 0 && co != paraNum - 1) {
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
                } else if (i == 0 && co != paraNum - 1) {
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
                } else if (co == paraNum - 1 && i != 0) {
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
            }
        }
        return tableList;
    }

    @Override
    public List<ImagePO> getAllImages(MultipartFile file) throws IOException {
        return null;
    }

    @Override
    public List<TitlePO> getAllTitle(MultipartFile file) throws IOException {
        return null;
    }

    @Override
    public ParagraphPO getParagraphText(Paragraph paragraph) {
        return null;
    }

    @Override
    public TitlePO getParagraphFormat(Paragraph paragraph) {
        return null;
    }

    @Override
    public FontPO getParagraphFontFormat(Paragraph paragraph) {
        return null;
    }

    @Override
    public List<ParagraphPO> getParagraphByTitle(MultipartFile file, int paragraphId) {
        return null;
    }

    @Override
    public List<ImagePO> getImagesByTitle(MultipartFile file, int paragraphId) {
        return null;
    }

    @Override
    public List<TablePO> getTablesByTitle(MultipartFile file, int paragraphId) {
        return null;
    }

    @Override
    public void delete() {

    }

}
