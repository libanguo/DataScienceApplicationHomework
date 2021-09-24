package com.example.application.serviceImpl;

import com.example.application.PO.*;
import com.example.application.service.PdfService;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

@Service
public class PdfServiceImpl implements PdfService {
    @Override
    public List<ParagraphPO> getAllParagraphs(MultipartFile file) throws IOException {
//        InputStream is=file.getInputStream();
//        PdfDocument pdfDocument=new PdfDocument();
//        pdfDocument.loadFromStream(is);
//        int pageCount=pdfDocument.getPages().getCount();
//        StringBuilder sb = new StringBuilder();
//        InputStream is= file.getInputStream();
//        ParagraphAbsorber paragraphAbsorber=new ParagraphAbsorber();
//        paragraphAbsorber.visit(new Document(is));
//        for (PageMarkup markup : paragraphAbsorber.getPageMarkups()) {
//            int i = 1;
//            for (MarkupSection section : markup.getSections()) {
//                int j = 1;
//
//                for (MarkupParagraph paragraph : section.getParagraphs()) {
//                    StringBuilder paragraphText = new StringBuilder();
//                    for (java.util.List<TextFragment> line : paragraph.getLines()) {
//                        for (TextFragment fragment : line) {
//                            paragraphText.append(fragment.getText());
//                        }
//                        paragraphText.append("\r\n");
//                    }
//                    paragraphText.append("\r\n");
//
//                    System.out.println("Paragraph "+j+" of section "+ i + " on page"+ ":"+markup.getNumber());
//                    System.out.println(paragraphText.toString());
//
//                    j++;
//                }
//                i++;
//            }
//        }
        return null;
    }

    @Override
    public List<TablePO> getAllTables(MultipartFile file) throws IOException {
        return null;
    }

    @Override
    public List<ImagePO> getAllImages(MultipartFile file) throws IOException {
        return null;
    }

    @Override
    public List<TitlePO> getAllTitles(MultipartFile file) throws IOException {
        return null;
    }

    @Override
    public ParagraphPO getParagraphText(Paragraph paragraph, int id) {
        return null;
    }

    @Override
    public ParagraphFormatPO getParagraphFormat(Paragraph paragraph, int id) {
        return null;
    }

    @Override
    public FontPO getParagraphFontFormat(Paragraph paragraph) {
        return null;
    }

    @Override
    public List<ParagraphPO> getParagraphByTitle(MultipartFile file, int paragraphId) throws IOException {
        return null;
    }

    @Override
    public List<ImagePO> getImagesByTitle(MultipartFile file, int paragraphId) throws IOException {
        return null;
    }

    @Override
    public List<TablePO> getTablesByTitle(MultipartFile file, int paragraphId) throws IOException {
        return null;
    }
}
