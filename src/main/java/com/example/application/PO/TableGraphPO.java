package com.example.application.PO;
import lombok.Getter;
import lombok.Setter;

@Setter
@Getter

public class TableGraphPO {
    private String tableTextContent;
    private int paragraphId;

    public TableGraphPO() {

    }

    public TableGraphPO(int paragraphId,String tableTextContent){
        this.paragraphId=paragraphId;
        this.tableTextContent=tableTextContent;
    }
}
