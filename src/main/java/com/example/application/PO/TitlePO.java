package com.example.application.PO;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter

public class TitlePO {
    private String paragraphText;
    private int paragraphId;
    private int lineSpacing;
    private int indentFromLeft;
    private int indentFromRight;
    private int firstLineIndent;
    private int lvl;
}
