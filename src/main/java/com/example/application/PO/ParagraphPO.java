package com.example.application.PO;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter

public class ParagraphPO {
    private String paragraphText;
    private int paragraphId;
    private int fontSize;
    private String fontName;
    private Boolean isBold;
    private Boolean isItalic;
    private Boolean isInTable;
    private int lvl;
    private int lineSpacing;
    private int fontAlignment;
    private Boolean isTableRowEnd;
    private int indentFromLeft;
    private int indentFromRight;
    private int firstLineIndent;
    public ParagraphPO(){

    }
}
