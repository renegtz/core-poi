package mx.infotec.dads.arq.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.Collator;
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

        } catch (IOException ex) {
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
            Object[] s;
            ImplementExcel excel = new ImplementExcel("C:\\Users\\rene\\Documents\\Actualizacion3.xlsx");
            List<Object[]> lst = excel.GetFilasColumnas(excel.getSheet().get(2));
            String Pcausas[] = {"Cambio de equipo", "Comisariato", "Carga", "Espera de equipo", "Evento ocacional", "Falta de certificado de aeronave", "Mantenimiento aeronaves", "Operaciones aerolínea", "Procedimiento de seguridad", "Rampa aerolínea", "Repercuciones", "Tráfico/documentación", "Tripulaciones"};
            String nopCausas[] = {"Accidente por un tercero", "Aerocares", "Aplicación de control de flujo", "Arco detector rayos X", "Autoridades", "Bloqueo carretera", "Cierre de aeropuerto", "Combustibles", "Control de flujo", "Control de flujo AICM", "Control de flujo LAX", "Control de flujo SENEAM", "Demora en ruta", "Emergencia médica", "Espera de equipo", "Evento ocacional", "Handler", "Impacto de ave", "Inauguración", "Incidente por un tercero", "Infraestrutura aeroportuaria", "Meteorología", "Ocacionada en su origen", "Otros", "Pasajero enfermo", "Pasajeros especiales", "Pasillos", "Repercuciones en ruta", "Repercuciones por un tercero", "Saturación de servicios", "Servicios de apoyo en tierra", "Suministro combustible TDE", "Visita papal", "Visita precidencial"};

            int mes;
            ArrayList causas = new ArrayList();
            ArrayList puntabilidad = new ArrayList();
            List<Object> valor = new ArrayList();
            List<Object[]> valor2 = new ArrayList();
            String nombre = "";
            int imputable = 1;

            for (int i = 0; i < lst.size(); i++) {
                s = lst.get(i);
                for (int x = 0; x < s.length; x++) {
                    if (i == 2) {
                        nombre = s[1].toString();
                    }
                    if (i >= 5 && x == 0 && !s[0].toString().equals("error") && !s[0].toString().equals("Total general")) {
                        if (s[0].toString().equals("No Imputable")) {
                            imputable = 0;

                        }
                        if (!s[0].toString().equals("No Imputable")) {
                            if (imputable == 1) {
                                s[0] = s[0].toString().substring(0, s[0].toString().length() - 1);
                            }
                            causas.add(s[0].toString());
                            puntabilidad.add(imputable);
                        }

                    }
                    if (i >= 5 && !s[0].toString().equals("No Imputable") && !s[0].toString().equals("Total general")) {
                        if (x >= 1 && i >= 5 && !s[0].toString().equals("error")) {
                            Object value = s[x];
                            valor.add(value);
                            if (x == s.length - 1) {

                                valor2.add(valor.toArray());
                                valor = new ArrayList();
                            }
                        }
                    }

                    System.out.print(x + " " + i + "  " + s[x] + "  ");
                }
                System.out.println();
            }

            System.out.println("El nombre de la aerolinea es " + nombre);
            System.out.println(causas);
            System.out.println(puntabilidad);
            System.out.println(valor);
//            System.out.println(valor2);
//             System.out.println(valor2.size());
            Object[] v;

//            System.out.println(v[0]);
            v = valor2.get(0);
            System.out.println("tamaño " + v.length);
            Collator comparador = Collator.getInstance();
            comparador.setStrength(Collator.PRIMARY);
// Estas dos cadenas son iguales
            System.out.println(comparador.compare("Hóla", "hola"));
            int mess = 1;
            for (int i = 0; i < v.length; i++) {

                for (int j = 0; j < Pcausas.length; j++) {
                    boolean paso = true;
                    for (int k = 0; k < causas.size(); k++) {
                        if (comparador.compare(Pcausas[j], causas.get(k)) == 0) {

                            paso = false;
                        }
                    }
                    if (paso == true) {
                        System.out.println("('" + nombre + "', '" + Pcausas[j] + "', 1, 0, " + mess + "),");

                    }
                }

                for (int x = 0; x < causas.size(); x++) {

                    if (x == causas.size() - 1) {
                        System.out.print("  ('" + nombre + "', '" + causas.get(x) + "', " + puntabilidad.get(x) + ", " + " " + v[i].toString() + ", " + mess + "),");
 System.out.println("");

                    } else {
                        System.out.print("  ('" + nombre + "', '" + causas.get(x) + "', " + puntabilidad.get(x) + ", " + " " + v[i].toString() + ", " + mess + "),");

                    }

                    System.out.println("");
                }
                for (int j = 0; j < nopCausas.length; j++) {
                    boolean paso = true;
                    for (int k = 0; k < causas.size(); k++) {
                        if (comparador.compare(nopCausas[j], causas.get(k)) == 0) {

                            paso = false;
                        }
                    }
                    if (paso == true) {
                        if (j == nopCausas.length - 1) {
                            System.out.println("('" + nombre + "', '" + nopCausas[j] + "', 0, 0, " + mess + ")");
                            
                        } else {

                            System.out.println("('" + nombre + "', '" + nopCausas[j] + "', 0, 0, " + mess + "),");
                        }

                    }
                }
                System.out.println("");
                System.out.println("");
                System.out.println("");
                System.out.println("");
                mess = mess + 1;

            }

        } catch (ExcelException ex) {
            Logger.getLogger(ImplementExcel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
