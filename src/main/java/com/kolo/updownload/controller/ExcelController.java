package com.kolo.updownload.controller;

import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.kolo.updownload.utils.ExcelUtils;
import com.kolo.updownload.utils.PoiExcelUtils;
import com.kolo.updownload.vo.BaseCodeVo;
import com.kolo.updownload.vo.BaseDataItemVo;
import com.kolo.updownload.vo.TargetDataItemVo;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.beans.IntrospectionException;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

@Controller
public class ExcelController {
    //跳转到上传文件的页面
    @RequestMapping(value = "/uploadimg", method = RequestMethod.GET)
    public String goUploadImg() {
        //跳转到 templates 目录下的 uploadimg.html
        return null;
    }
    @PostMapping("/upload")
    @ResponseBody
    public String upload(@RequestParam("file") MultipartFile file,
                         HttpServletRequest request) throws Exception {
        String contentType = file.getContentType();
        String fileName = file.getOriginalFilename();  //获取上传文件原名
        ImportParams params = new ImportParams();
        params.setTitleRows(2);
        params.setHeadRows(1);
        params.setSheetNum(5);
        List<BaseDataItemVo> baseDataItemVoList = PoiExcelUtils.importExcel(file.getInputStream(), 4, 2, 1, BaseDataItemVo.class);
        List<BaseCodeVo> baseCodeVoList = PoiExcelUtils.importExcel(file.getInputStream(), 5, 1, 2, BaseCodeVo.class);
        List<TargetDataItemVo> targetDataItemVoList = PoiExcelUtils.importExcel(file.getInputStream(), 6, 7, 2, TargetDataItemVo.class);
        //List<BaseDataItemVo> personVoList = ExcelUtils.importExcel(file.getInputStream(), BaseDataItemVo.class, params);
        return "success";
    }

    @RequestMapping("/exportExcel/{name}")
    public void mecExportExcel(@PathVariable(name = "name") String name, HttpServletResponse response) throws Exception {



        XSSFWorkbook workbook = null;
        if (name.equals("baseDataIteam")){
            List<BaseDataItemVo> list = test1();
            workbook = PoiExcelUtils.createWorkbook("dataItemModel.xlsx", list, BaseDataItemVo.class);
        }else if (name.equals("baseCode")){
            List<BaseCodeVo> list = test2();
            workbook = PoiExcelUtils.createWorkbook("dataCodeModel.xlsx", list, BaseCodeVo.class);
        }else if(name.equals("targetDataIteam")){
            List<TargetDataItemVo> list = test3();
            workbook = PoiExcelUtils.createWorkbook("targetDataItemModel.xlsx", list, TargetDataItemVo.class);
        }
        response.reset();
        response.setContentType("application/msexcel");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition",  "attchment;filename=" + new String("123.xlsx".getBytes()) );

        OutputStream out = response.getOutputStream();
        workbook.write(out);
        out.flush();
    }

    private List<BaseDataItemVo> test1() throws Exception {
        List<BaseDataItemVo> list;
        String filePath = PoiExcelUtils.class.getClassLoader().getResource("static" + "/" + "2.xlsx").getPath();
        list = PoiExcelUtils.importExcel(new FileInputStream(new File(filePath)), 4, 2, 1, BaseDataItemVo.class);
        return list;
    }

    private List<BaseCodeVo> test2() throws Exception {
        List<BaseCodeVo> list;
        String filePath = PoiExcelUtils.class.getClassLoader().getResource("static" + "/" + "2.xlsx").getPath();
        list = PoiExcelUtils.importExcel(new FileInputStream(new File(filePath)), 5, 1, 2, BaseCodeVo.class);
        return list;
    }

    private List<TargetDataItemVo> test3() throws Exception {
        List<TargetDataItemVo> list;
        String filePath = PoiExcelUtils.class.getClassLoader().getResource("static" + "/" + "2.xlsx").getPath();
        list = PoiExcelUtils.importExcel(new FileInputStream(new File(filePath)), 6, 7, 2, TargetDataItemVo.class);
        return list;
    }

}
