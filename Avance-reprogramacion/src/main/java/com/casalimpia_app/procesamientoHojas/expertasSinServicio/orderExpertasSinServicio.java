package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class orderExpertasSinServicio {
    public static void reorganizeExcelExpertasSinServicio(Workbook wb) throws Exception {
        // Obtener la hoja "Expertas Sin Servicio"
        Sheet sheet = wb.getSheet("Expertas Sin Servicio");
        if (sheet == null) {
            throw new Exception("La hoja 'Expertas Sin Servicio' no existe.");
        }

        // Leer las filas, excluyendo la cabecera
        List<Row> rows = new ArrayList<>();
        //int headerIndex = sheet.getFirstRowNum(); // Suponemos que la primera fila es la cabecera
        int lastRowIndex = sheet.getLastRowNum();

        // Almacenar las filas (sin incluir la cabecera)
        for (int i = 0; i <= lastRowIndex; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                rows.add(row);
            }
        }

        // Ordenar las filas alfabéticamente basándose en la columna E (índice 4)
        rows.sort(Comparator.comparing(row -> {
            Cell cell = row.getCell(4); // Columna E
            return cell != null ? cell.getStringCellValue() : "";
        }));

        // Crear una nueva hoja para escribir los datos reorganizados
        Sheet sortedSheet = wb.createSheet("Expertas Sin Servicio Ordenado");
        // Copiar la cabecera
        /*
        Row header = sheet.getRow(headerIndex);
        if (header != null) {
            copyRow(header, sortedSheet.createRow(0));
        }
         */
        // Escribir las filas ordenadas en la nueva hoja
        int rowIndex = 0;
        for (Row row : rows) {
            copyRow(row, sortedSheet.createRow(rowIndex++));
        }

        // Opcional: Eliminar la hoja original (si se desea)
        int sheetIndex = wb.getSheetIndex(sheet);
        wb.removeSheetAt(sheetIndex);

        // Renombrar la nueva hoja
        wb.setSheetName(wb.getSheetIndex(sortedSheet), "Expertas Sin Servicio");
        
    }

    private static void copyRow(Row sourceRow, Row targetRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(i);
                targetCell.setCellStyle(sourceCell.getCellStyle());
                switch (sourceCell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        targetCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        targetCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        break;
                }
            }
        }
    }
    /*
    public static void main(String[] args) throws Exception {
        Workbook wb = new XSSFWorkbook(); // Cargar el workbook
        // Llama a la función reorganizeExcelExpertasSinServicio aquí con tu workbook cargado
    }
    */
}
