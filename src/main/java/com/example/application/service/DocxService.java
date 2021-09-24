package com.example.application.service;

import com.example.application.PO.*;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public interface DocxService {
    public void wordParse(MultipartFile file, HashMap<String, List<ParagraphPO>> paragraphHashMap,HashMap<String, List<TablePO>> tableHashMap,HashMap<String, List<ImagePO>> imageHashMap,HashMap<String, List<TitlePO>> titleHashMap,HashMap<String,List<FontPO>> fontHashMap,String token) throws IOException;

    public ParagraphFormatPO getParagraphFormat(ParagraphPO paragraph);



}
