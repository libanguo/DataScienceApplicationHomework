package com.example.application.serviceImpl;

import com.example.application.PO.*;
import com.example.application.service.DocxService;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
public class DocxServiceImpl implements DocxService {
    @Override
    public List<ParagraphPO> getAllParagraphs(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        XWPFDocument doc = new XWPFDocument(is);
        List<ParagraphPO> paragraphList = new ArrayList<>();
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (XWPFParagraph paragraph : paras) {
            ParagraphPO paragraphPO = new ParagraphPO();
            XWPFRun xwpfRun = paragraph.getRuns().get(0);
            paragraphPO.setParagraphId(paragraph.getNumID().intValue());
            paragraphPO.setParagraphText(paragraphPO.getParagraphText());
            paragraphPO.setFontSize(xwpfRun.getFontSize());
            paragraphPO.setFontName(xwpfRun.getFontName());
            paragraphPO.setIsBold(xwpfRun.isBold());
            paragraphPO.setIsItalic(xwpfRun.isItalic());
            paragraphPO.setIsInTable(false);
            //TODO 存疑
            paragraphPO.setLvl(Integer.parseInt(paragraph.getStyle()));
            paragraphPO.setLineSpacing(paragraph.getSpacingLineRule().getValue());
            paragraphPO.setFontAlignment(paragraph.getFontAlignment());
            paragraphPO.setIsTableRowEnd(false);
            paragraphPO.setIndentFromLeft(paragraph.getIndentFromLeft());
            paragraphPO.setIndentFromRight(paragraph.getIndentFromRight());
            paragraphList.add(paragraphPO);
        }
        return paragraphList;
    }

    @Override
    //TODO
    public List<TablePO> getAllTables(MultipartFile file) throws IOException {
        return null;
    }

    @Override
    public List<ImagePO> getAllImages(MultipartFile file) throws IOException {
        List<ImagePO> imagePOList = new ArrayList<>();
        InputStream is = file.getInputStream();
        XWPFDocument doc = new XWPFDocument(is);
        List<XWPFPictureData> picList = doc.getAllPictures();
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (XWPFParagraph paragraph : paras) {
            XWPFRun xwpfRun = paragraph.getRuns().get(0);
            for (XWPFPicture pic : xwpfRun.getEmbeddedPictures()) {
                ImagePO imagePO = new ImagePO();
                imagePO.setWidth(pic.getWidth());
                imagePO.setHeight(pic.getDepth());
                imagePO.setTextBefore(pic.getDescription());
                imagePO.setTextAfter(pic.getDescription());
                imagePO.setSuggestFileExtension(pic.getPictureData().suggestFileExtension());
                imagePO.setFilename(pic.getPictureData().getFileName());
                imagePO.setBase64Content(pic.getPictureData().getData());
                imagePOList.add(imagePO);
            }
        }
        return imagePOList;
    }

    @Override
    public List<TitlePO> getAllTitles(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        XWPFDocument doc = new XWPFDocument(is);
        List<XWPFParagraph> paras = doc.getParagraphs();
        List<TitlePO> titlelist = new ArrayList<>();
        for (XWPFParagraph graph : paras) {
            String style = graph.getStyle();
            if (style.compareTo("9") < 0) {
                TitlePO title = new TitlePO();
                title.setParagraphText(graph.getParagraphText());
                title.setParagraphId(graph.getNumID().intValue());
                title.setIndentFromLeft(graph.getIndentFromLeft());
                title.setIndentFromRight(graph.getIndentFromRight());
                title.setFirstLineIndent(graph.getFirstLineIndent());
                title.setLvl(Integer.parseInt(graph.getStyle()));
                titlelist.add(title);
            }
        }
        return titlelist;
    }

    @Override
    public ParagraphPO getParagraphText(XWPFParagraph paragraph, int id) {
        ParagraphPO paragraphPO = new ParagraphPO();
        XWPFRun xwpfRun = paragraph.getRuns().get(0);
        paragraphPO.setParagraphText(paragraph.getParagraphText());
        paragraphPO.setParagraphId(Integer.parseInt(paragraph.getStyle()));
        paragraphPO.setFontSize(xwpfRun.getFontSize());
        paragraphPO.setFontName(xwpfRun.getFontName());
        paragraphPO.setIsBold(xwpfRun.isBold());
        paragraphPO.setIsItalic(xwpfRun.isItalic());
        paragraphPO.setIsInTable(false);
        //TODO 存疑
        paragraphPO.setLvl(Integer.parseInt(paragraph.getStyle()));
        paragraphPO.setLineSpacing(paragraph.getSpacingLineRule().getValue());
        paragraphPO.setFontAlignment(paragraph.getFontAlignment());
        paragraphPO.setIsTableRowEnd(false);
        paragraphPO.setIndentFromLeft(paragraph.getIndentFromLeft());
        paragraphPO.setIndentFromRight(paragraph.getIndentFromRight());
        return paragraphPO;
    }

    @Override
    public ParagraphFormatPO getParagraphFormat(XWPFParagraph paragraph, int id) {
        ParagraphFormatPO paragraphFormatPO = new ParagraphFormatPO();
        paragraphFormatPO.setLvl(Integer.parseInt(paragraph.getStyle()));
        paragraphFormatPO.setIndentFromLeft(paragraph.getIndentFromLeft());
        paragraphFormatPO.setIndentFromRight(paragraph.getIndentFromRight());
        paragraphFormatPO.setFirstLineIndent(paragraph.getFirstLineIndent());
        paragraphFormatPO.setLineSpacing(paragraph.getSpacingLineRule().getValue());
        return paragraphFormatPO;
    }

    @Override
    public FontPO getParagraphFontFormat(XWPFParagraph paragraph) {
        FontPO fontPO = new FontPO();
        XWPFRun xwpfRun = paragraph.getRuns().get(0);
        fontPO.setColor(Integer.parseInt(xwpfRun.getColor()));
        fontPO.setFontSize(xwpfRun.getFontSize());
        fontPO.setFontName(xwpfRun.getFontName());
        fontPO.setIsBold(xwpfRun.isBold());
        fontPO.setIsItalic(xwpfRun.isItalic());
        fontPO.setFontAlignment(paragraph.getFontAlignment());
        return fontPO;
    }

    @Override
    //TODO
    public List<ParagraphPO> getParagraphByTitle(MultipartFile file, int paragraphId) throws IOException {
        return null;
    }

    @Override
    //TODO
    public List<ImagePO> getImagesByTitle(MultipartFile file, int paragraphId) throws IOException {
        return null;
    }

    @Override
    //TODO
    public List<TablePO> getTablesByTitle(MultipartFile file, int paragraphId) throws IOException {
        return null;
    }
}
