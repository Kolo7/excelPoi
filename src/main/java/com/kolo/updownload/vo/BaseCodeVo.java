package com.kolo.updownload.vo;

import lombok.Data;

import java.io.Serializable;

@Data
public class BaseCodeVo implements Serializable {
    private String codeNumber;

    private String codeName;

    private String codeValue;

    private String codeItemName;

    private String codeDesc;

    private String codeRule;

    private String remarks;
}
