package com.example.application.PO;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter

public class ImagePO {
    private int paragraphBefore;
    private String textBefore;
    private String textAfter;
    private Double height;
    private Double width;
    private String suggestFileExtension;
    private byte[] base64Content;
    private String filename;
}
