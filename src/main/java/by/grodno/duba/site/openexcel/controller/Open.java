package by.grodno.duba.site.openexcel.controller;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;


@Controller
public class Open {
    @GetMapping("/open")
    public static void main(String[] args, Model model) throws IOException {
        String excelFilePath = "Test.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();

        model.addAttribute("cell", iterator);



       while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            model.addAttribute("cellIterator", cellIterator);

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                }}
/*
                     System.out.print(cell.getStringCellValue());

//                switch (cell.getCellType()) {
//                    case CellType.STRING:
//                        System.out.print(cell.getStringCellValue());
//                        break;
//                    case Cell.CELL_TYPE_BOOLEAN:
//                        System.out.print(cell.getBooleanCellValue());
//                        break;
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.print(cell.getNumericCellValue());
//                        break;
//                }
                System.out.print(" - ");
            }
            System.out.println();
        }

        workbook.close();
        inputStream.close();
*/
    }
}
