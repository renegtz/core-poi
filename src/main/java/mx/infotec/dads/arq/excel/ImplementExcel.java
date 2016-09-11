package mx.infotec.dads.arq.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import mx.infotec.dads.arq.excel.exception.ExcelException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;

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

    private final String path;

    public ImplementExcel(String path) {
        this.path = path;
    }

    public List<XSSFSheet> getSheet() throws ExcelException {
        InputStream excelFileToRead = null;
        List<XSSFSheet> lst = new ArrayList<>();
        try {
            excelFileToRead = new FileInputStream(this.path);
            XSSFWorkbook wb = new XSSFWorkbook(excelFileToRead);
            int numberOfSheet = wb.getNumberOfSheets();
            for (int i = 0; i < numberOfSheet; i++) {
                lst.add(wb.getSheetAt(i));
            }

        } catch ( IOException ex) {
            throw new ExcelException("Error a obtener los libros del excel ", ex);
        } finally {
            try {
                excelFileToRead.close();
            } catch (IOException ex) {
                Logger.getLogger(ImplementExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        return lst;
    }

    public String getStringCellValue(XSSFCell cell) {
        String value;
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue() + " ";
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                Double n = cell.getNumericCellValue();
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue().toString();
                } else {
                    value = n + " ";
                }
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue() + " ";
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                value = cell.getErrorCellString();
                break;
            default:
                value = "error";
                break;
        }

        return value;
    }

    public Object getCellValue(XSSFCell cell) {
        Object value;
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                value = cell.getStringCellValue();
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue();
                } else {
                    value = cell.getNumericCellValue();
                }
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                value = cell.getErrorCellString();
                break;
            default:
                value = "error";
                break;
        }

        return value;
    }

    public List<Object> getFila(XSSFRow row) {
        List<Object> lst = new ArrayList<>();
        Iterator cells = row.cellIterator();
        while (cells.hasNext()) {
            XSSFCell cell = (XSSFCell) cells.next();
            lst.add(getCellValue(cell));
        }
        return lst;
    }
    
    public Object[] getFilaArreglo(XSSFRow row) {
        List<Object> lst = new ArrayList<>();
        Iterator cells = row.cellIterator();
        while (cells.hasNext()) {
            XSSFCell cell = (XSSFCell) cells.next();
            lst.add(getCellValue(cell));
        }
        return lst.toArray();
    }
    
    public List<String> getFilaString(XSSFRow row) {
        List<String> lst = new ArrayList<>();
        Iterator cells = row.cellIterator();
        while (cells.hasNext()) {
            XSSFCell cell = (XSSFCell) cells.next();
            lst.add(getStringCellValue(cell));
        }
        return lst;
    }

    @Override
    public List<Object[]> GetFilasColumnas(XSSFSheet sheet) throws ExcelException {
        List<Object[]> lst = new ArrayList<>();
        try {

            XSSFRow row;

            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                row = (XSSFRow) rows.next();
                lst.add(getFilaArreglo(row));
            }
        } catch (Exception e) {
            throw new ExcelException("Error al procesar excel", e);
        }

        return lst;
    }

    public static void main(String[] args) {
        try {
            ImplementExcel excel = new ImplementExcel("/home/abel/Descargas/excel.xlsx");
            List<Object[]> lst = excel.GetFilasColumnas(excel.getSheet().get(0));
            System.out.println("numeros de filas "+lst.size());
        } catch (ExcelException ex) {
            Logger.getLogger(ImplementExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
