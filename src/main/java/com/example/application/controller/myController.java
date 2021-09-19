package com.example.application.controller;

import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiResponse;
import io.swagger.annotations.ApiResponses;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
@CrossOrigin
public class myController {

    @ApiOperation(value = "上传文件")
    @ApiResponses({ @ApiResponse(code = 200, message = "success", response =String.class) })
    @PostMapping("/load_file")
    public String addUser(String file,String fileName) {
        System.out.println(file);
        System.out.println(fileName);
        return "abc";
    }
}
