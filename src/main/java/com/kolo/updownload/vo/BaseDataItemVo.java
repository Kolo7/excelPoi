package com.kolo.updownload.vo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import java.io.Serializable;

@Data
public class BaseDataItemVo implements Serializable {
    @Excel(name = "序号", orderNum = "1")
    private String no;

    @Excel(name = "一级数据主题", orderNum = "2")
    private String oneLevelDataTheme;

    @Excel(name = "二级数据主题", orderNum = "3")
    private String twoLevelDataTheme;

    @Excel(name = "三级数据主题", orderNum = "4")
    private String threeLevelDataTheme;


    @Excel(name = "实体", orderNum = "5")
    private String entity;

    @Excel(name = "数据项编号", orderNum = "6")
    private String dataItemNubmer;

    @Excel(name = "数据项中文名称", orderNum = "7")
    private String dataItemChineseName;

    @Excel(name = "数据英文名称", orderNum = "8")
    private String dataEnglishName;

    @Excel(name = "数据项常用名称", orderNum = "9")
    private String dateItemOftenName;

    @Excel(name = "代码名称", orderNum = "10")
    private String codeName;

    @Excel(name = "业务定义", orderNum = "11")
    private String businessDefine;

    @Excel(name = "执行标准", orderNum = "12")
    private String executeStandard;

    @Excel(name = "定义依据", orderNum = "13")
    private String defineBasis;

    @Excel(name = "对象分类", orderNum = "14")
    private String objectClassify;

    @Excel(name = "数据类型", orderNum = "15")
    private String dateType;

    @Excel(name = "数据长度", orderNum = "16")
    private String dateLength;

    @Excel(name = "数据定义部门", orderNum = "17")
    private String dateDefineDepartment;

    @Excel(name = "数据来源部门", orderNum = "18")
    private String dateSourceDepartment;

    @Excel(name = "备注", orderNum = "19")
    private String remarks;

    @Excel(name = "系统名称", orderNum = "20")
    private String systemName;

    @Excel(name = "数据库英文名称", orderNum = "21")
    private String databaseEnglishName;

    @Excel(name = "Schema英文名称", orderNum = "22")
    private String schemaEnglishName;

    @Excel(name = "表名称", orderNum = "23", type = 3)
    private String tableName;

    @Excel(name = "表字段名称", orderNum = "24")
    private String tableFieldName;
}
