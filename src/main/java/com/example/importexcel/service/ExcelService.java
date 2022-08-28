package com.example.importexcel.service;

import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface ExcelService  {
    String excelImport(MultipartFile file);
}
