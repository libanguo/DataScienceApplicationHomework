package com.example.application.PO;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter

public class ImagePO {
    private String textBefore;
    private String textAfter;
    private int height;
    private int width;
    private String suggestFileExtension;
    private byte[] base64Content;
    private String filename;
}
