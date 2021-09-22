package com.example.application.serviceImpl;


import com.example.application.PO.ParagraphPO;
import com.example.application.service.DocService;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
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
            paragraphPO.setLineSpacing(paragraph.getLineSpacing());
            paragraphPO.setIsInTable(paragraph.isInTable());
            paragraphPO.setLvl(paragraph.getLvl());
            paragraphPO.setIsTableRowEnd(paragraph.isTableRowEnd());
            paragraphList.add(paragraphPO);
        }
        return paragraphList;
    }
}
