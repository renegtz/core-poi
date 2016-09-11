/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package mx.infotec.dads.arq.excel;

import java.util.List;
import mx.infotec.dads.arq.excel.exception.ExcelException;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 *
 * @author abel
 */
public interface Excel {
    
    List<Object[]> GetFilasColumnas(XSSFSheet sheet) throws ExcelException;
    
}
