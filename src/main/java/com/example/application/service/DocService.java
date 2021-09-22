package com.example.application.service;


import com.example.application.PO.ParagraphPO;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface DocService {
    public List<ParagraphPO> getAllParagraphs(MultipartFile file) throws IOException;

}
