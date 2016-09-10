package com.mycompany.imp.poi;


import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.mycompany.core.poi.Excel;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author rene
 */
public class ImplementExcel implements Excel {

    @Override
    public String GetTxt(String Patch) {

        String txt = "";
        try {
            InputStream ExcelFileToRead = new FileInputStream(Patch);
            XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;

            Iterator rows = sheet.rowIterator();

            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                Iterator cells = row.cellIterator();
                while (cells.hasNext()) {
                    cell = (XSSFCell) cells.next();

                    if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        System.out.print(cell.getStringCellValue() + " ");
                    } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                         txt+=cell.getNumericCellValue()+"\n";
                    } else {
                    }
                }
            }
        } catch (Exception e) {
        }

        return txt;
    }
    public static void main(String[] args) {
        ImplementExcel excel = new ImplementExcel();
        System.out.println(excel.GetTxt("C:\\datos\\excel.xlsx"));
    }

}
