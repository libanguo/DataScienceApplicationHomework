package com.example.application.VO;

import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiParam;


@ApiModel(value="RestMessage", description="RestMessage desc")
public class ResponseVO {

    /**
     * 调用是否成功
     */
    @ApiParam(value="message")
    @JsonSerialize(include= JsonSerialize.Inclusion.NON_EMPTY)
    private int code=0;

    /**
     * 返回的提示信息
     */
    @JsonSerialize(include= JsonSerialize.Inclusion.NON_EMPTY)
    private String msg="success";

    /**
     * 内容
     */
    @JsonSerialize(include= JsonSerialize.Inclusion.NON_EMPTY)
    private Object data;

    public static ResponseVO buildSuccess(Object data){
        ResponseVO responseVO=new ResponseVO();
        responseVO.data=data;
        return responseVO;
    }

    public static ResponseVO buildFailure(String message){
        ResponseVO responseVO=new ResponseVO();
        responseVO.msg=message;
        return responseVO;
    }
}
