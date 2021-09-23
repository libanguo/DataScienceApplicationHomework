package com.example.application.service;

import com.example.application.PO.*;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface DocxService {
    public List<ParagraphPO> getAllParagraphs(MultipartFile file) throws IOException;

    public List<TablePO> getAllTables(MultipartFile file) throws IOException;

    public List<ImagePO> getAllImages(MultipartFile file) throws IOException;

    public List<TitlePO> getAllTitles(MultipartFile file) throws IOException;

    public ParagraphPO getParagraphText(XWPFParagraph paragraph, int id);

    public ParagraphFormatPO getParagraphFormat(XWPFParagraph paragraph, int id);

    public FontPO getParagraphFontFormat(XWPFParagraph paragraph);

    public List<ParagraphPO> getParagraphByTitle(MultipartFile file, int paragraphId) throws IOException;

    public List<ImagePO> getImagesByTitle(MultipartFile file, int paragraphId) throws IOException;

    public List<TablePO> getTablesByTitle(MultipartFile file, int paragraphId) throws IOException;

}
