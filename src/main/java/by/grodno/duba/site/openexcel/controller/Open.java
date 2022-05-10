package by.grodno.duba.site.openexcel.controller;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import by.grodno.duba.site.openexcel.poi.excel.ExcelPOIHelper;


@Controller
public class Open {


    private ExcelPOIHelper excelPOIHelper;
    private static String FILE_NAME = "Test.xlsx";
    private String fileLocation;


    @GetMapping("/open")
    public String main(Model model) throws IOException {
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        fileLocation = path.substring(0, path.length() - 1) + FILE_NAME;

        excelPOIHelper = new ExcelPOIHelper();
        excelPOIHelper.writeExcel();



        String excelFilePath = "Test.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Object cellValue = new Object();

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();


        if (excelFilePath != null) {
            if (excelFilePath.endsWith(".xlsx") || excelFilePath.endsWith(".xls")) {


                Map<Integer, List<String>> data
                        = excelPOIHelper.readExcel(fileLocation);
                model.addAttribute("data", data);
            } else {
                model.addAttribute("message", "Not a valid excel file!");
            }
        } else {
            model.addAttribute("message", "File missing! Please upload an excel file.");
        }
        return "open";
    }


}


