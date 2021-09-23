package com.example.application.controller;

import com.example.application.VO.ResponseVO;
import com.example.application.service.DocService;
import com.example.application.service.DocxService;
import com.example.application.service.PdfService;
import com.example.application.tool.Tool;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.HashMap;
import java.util.List;

@RestController
@CrossOrigin
public class MyController {

    @Autowired
    private DocService docService;

    @Autowired
    private DocxService docxService;

    @Autowired
    private PdfService pdfService;

    HashMap<String, MultipartFile> hashMap = new HashMap<>();

    @ApiOperation(value = "上传文件")
    @PostMapping("/load_file")
    public ResponseVO loadFile(@RequestParam("file") MultipartFile file, @RequestParam("fileName") String fileName) throws IOException {
        String[] tmp = fileName.split("\\.");
        String type = tmp[tmp.length - 1];
        if (type.equals("doc") || type.equals("pdf") || type.equals("docx") || type.equals("wps")) {
            InputStream inputStream = file.getInputStream();
            StringBuilder sb = new StringBuilder();
            String line;
            BufferedReader br = new BufferedReader(new InputStreamReader(inputStream));
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
            String content = sb.toString();
            String token = Tool.SHA(content);
            hashMap.put(token, file);
            return ResponseVO.buildSuccess(token);
        } else {
            return ResponseVO.buildSuccess("上传文件类型无法解析");
        }
    }

    @ApiOperation(value = "根据Token获取文档内全部段落信息")
    @PostMapping("/word_parser/{token}/all_paragraphs")
    public ResponseVO getAllParagraphs(@PathVariable String token) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getAllParagraphs(file));
                case "docx":
                    return ResponseVO.buildSuccess(docxService.getAllParagraphs(file));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token获取文档内全部表格信息")
    @PostMapping("/word_parser/{token}/all_tables")
    public ResponseVO getAllTables(@PathVariable String token) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getAllTables(file));
                case "docx":
                    return ResponseVO.buildSuccess(docxService.getAllTables(file));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token获取文档内全部图片信息")
    @PostMapping("/word_parser/{token}/all_pics")
    public ResponseVO getAllImages(@PathVariable String token) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getAllImages(file));
                case "docx":
                    return ResponseVO.buildSuccess(docxService.getAllImages(file));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token获取文档内全部标题信息")
    @PostMapping("/word_parser/{token}/all_titles")
    public ResponseVO getAllTitles(@PathVariable String token) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getAllTitles(file));
                case "docx":
                    return ResponseVO.buildSuccess(docxService.getAllTitles(file));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token、段落id获取段落详细信息")
    @PostMapping("/word_parser/{token}/paragraph/{paragraph_id}")
    public ResponseVO getParagraphById(@PathVariable String token, int paragraph_id) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];

            switch (type) {
                case "doc":
                    InputStream is = file.getInputStream();
                    HWPFDocument doc = new HWPFDocument(is);
                    Range range = doc.getRange();
                    Paragraph paragraph = range.getParagraph(paragraph_id);
                    return ResponseVO.buildSuccess(docService.getParagraphText(paragraph, paragraph_id));
                case "docx":
                    InputStream iss = file.getInputStream();
                    XWPFDocument docx = new XWPFDocument(iss);
                    List<XWPFParagraph> paras = docx.getParagraphs();
                    XWPFParagraph para = paras.get(paragraph_id);
                    return ResponseVO.buildSuccess(docxService.getParagraphText(para, paragraph_id));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token、段落id获取段落格式")
    @PostMapping("/word_parser/{token}/paragraph/{paragraph_id}/paragraph_stype")
    public ResponseVO getParagraphFormatById(@PathVariable String token, int paragraph_id) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    InputStream is = file.getInputStream();
                    HWPFDocument doc = new HWPFDocument(is);
                    Range range = doc.getRange();
                    Paragraph paragraph = range.getParagraph(paragraph_id);
                    return ResponseVO.buildSuccess(docService.getParagraphFormat(paragraph, paragraph_id));
                case "docx":
                    InputStream iss = file.getInputStream();
                    XWPFDocument docx = new XWPFDocument(iss);
                    List<XWPFParagraph> paras = docx.getParagraphs();
                    XWPFParagraph para = paras.get(paragraph_id);
                    return ResponseVO.buildSuccess(docxService.getParagraphFormat(para, paragraph_id));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token、段落id获取段落详细字体格式")
    @PostMapping("/word_parser/{token}/paragraph/{paragraph_id}/font_stype")
    public ResponseVO getParagraphFontFormatById(@PathVariable String token, int paragraph_id) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    InputStream is = file.getInputStream();
                    HWPFDocument doc = new HWPFDocument(is);
                    Range range = doc.getRange();
                    Paragraph paragraph = range.getParagraph(paragraph_id);
                    return ResponseVO.buildSuccess(docService.getParagraphFontFormat(paragraph));
                case "docx":
                    InputStream iss = file.getInputStream();
                    XWPFDocument docx = new XWPFDocument(iss);
                    List<XWPFParagraph> paras = docx.getParagraphs();
                    XWPFParagraph para = paras.get(paragraph_id);
                    return ResponseVO.buildSuccess(docxService.getParagraphFontFormat(para));
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    //TODO
    @ApiOperation(value = "根据Token、段落id获取标题下全部段落信息")
    @PostMapping("/word_parser/{token}/title/{paragraph_id}/all_paragraphs")
    public ResponseVO getParagraphByTitle(@PathVariable String token, int paragraph_id) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getParagraphByTitle(file,paragraph_id));
                case "docx":
                    break;
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    //TODO
    @ApiOperation(value = "根据Token、段落id获取标题下全部图片信息")
    @PostMapping("/word_parser/{token}/title/{paragraph_id}/all_pics")
    public ResponseVO getImagesByTitle(@PathVariable String token, int paragraph_id) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getImagesByTitle(file,paragraph_id));
                case "docx":
                    break;
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    //TODO
    @ApiOperation(value = "根据Token、段落id获取标题下全部表格信息")
    @PostMapping("/word_parser/{token}/title/{paragraph_id}/all_tables")
    public ResponseVO getTablesByTitle(@PathVariable String token, int paragraph_id) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    return ResponseVO.buildSuccess(docService.getTablesByTitle(file,paragraph_id));
                case "docx":
                    break;
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法获取内容");
    }

    @ApiOperation(value = "根据Token释放资源")
    @PostMapping("/word_parser/{token}")
    public ResponseVO delete(@PathVariable String token) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file != null) {
            String fileName = file.getOriginalFilename();
            String[] tmp = fileName.split("\\.");
            String type = tmp[tmp.length - 1];
            switch (type) {
                case "doc":
                    hashMap.remove(token);
                    return ResponseVO.buildSuccess("释放成功");
                case "docx":
                    break;
                case "pdf":
                    break;
                default:
                    return ResponseVO.buildSuccess("");
            }

        }
        return ResponseVO.buildSuccess("无法释放");
    }
}
