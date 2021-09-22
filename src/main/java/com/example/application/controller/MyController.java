package com.example.application.controller;

import com.example.application.VO.ResponseVO;
import com.example.application.service.DocService;
import com.example.application.service.DocxService;
import com.example.application.service.PdfService;
import com.example.application.tool.Tool;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.HashMap;

@RestController
@CrossOrigin
public class MyController {

    @Autowired
    private DocService docService;

    @Autowired
    private DocxService docxService;

    @Autowired
    private PdfService pdfService;

    HashMap<String,MultipartFile> hashMap=new HashMap<>();
    @ApiOperation(value = "上传文件")
    @PostMapping("/load_file")
    public ResponseVO loadFile(@RequestParam("file") MultipartFile file, @RequestParam("fileName") String fileName) throws IOException {
        String[] tmp=fileName.split("\\.");
        String type=tmp[tmp.length-1];
        if(type.equals("doc")||type.equals("pdf")||type.equals("docx")||type.equals("wps")){
            InputStream inputStream=file.getInputStream();
            StringBuilder sb = new StringBuilder();
            String line;
            BufferedReader br = new BufferedReader(new InputStreamReader(inputStream));
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
            String content = sb.toString();
            String token= Tool.SHA(content);
            hashMap.put(token,file);
            return ResponseVO.buildSuccess(token);
        }
        else {
            return ResponseVO.buildSuccess("上传文件类型无法解析");
        }
    }

    @ApiOperation(value = "根据Token获取文档内全部段落信息")
    @PostMapping("/word_parser/{token}/all_paragraphs")
    public ResponseVO getAllParagraphs(@PathVariable String token) throws IOException {
        MultipartFile file = hashMap.get(token);
        if (file!=null){
            String fileName = file.getOriginalFilename();
            String[] tmp=fileName.split("\\.");
            String type=tmp[tmp.length-1];
            switch (type){
                case "doc":
                    return ResponseVO.buildSuccess(docService.getAllParagraphs(file));
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

}
