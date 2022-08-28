package com.example.importexcel.controller;

import com.example.importexcel.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import java.util.List;

@RestController
@RequestMapping("/excel")
public class ImportExcelController {

    @Autowired
    public ExcelService excelService;

    @GetMapping("/dept/list")
    public String queryAll(){
        System.out.println("1111111111111111111111111");
        return "1111111111111111111111111";
    }

    @RequestMapping("import-excel-page")
    public ModelAndView importExcelPage() {
        System.out.println("到这了");
        ModelAndView mv = new ModelAndView();
        mv.setViewName("import-excel/import-excel.html");
        return mv;
    }

    @RequestMapping("import")
    public String excelImport(@RequestParam("file") MultipartFile file) {
        return excelService.excelImport(file);
    }

    @RequestMapping("import-excel-el-page")
    public ModelAndView importExcelElPage() {
        ModelAndView mv = new ModelAndView();
        mv.setViewName("import-excel/import-excel.html");
        return mv;
    }
}
