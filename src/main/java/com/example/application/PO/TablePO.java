package com.example.application.PO;

import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hwpf.usermodel.Paragraph;

import java.util.List;

@Setter
@Getter

public class TablePO {
    private String textBefore;
    private String textAfter;
    private TableGraphPO paragraphBefore;
    private TableGraphPO paragraphAfter;
    private List<TableGraphPO> tableContent;
    public TablePO(){

    }
}
