/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.hoja1;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class intFormat {
    public static void convertirTextoANumeroHoja1(Workbook wb) throws IOException {

        Sheet ws = wb.getSheet("Hoja1");

        // Convertir columnas D (índice 3) y B (índice 2) 
        int[] columnas = {3, 2};

        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                for (int colIndex : columnas) {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null && cell.getCellType() == Cell.CELL_TYPE_STRING) {
                        String cellValue = cell.getStringCellValue();

                        // Verificar si el valor de la celda es numérico o contiene espacios al inicio o final
                        if (cellValue.matches("\\s*\\d+\\s*")) {
                            // Eliminar espacios en blanco y convertir a numérico
                            double numericValue = Double.parseDouble(cellValue.trim());
                            cell.setCellValue(numericValue);
                        }
                    }
                }
            }
        }
        System.out.println("Proceso completado. Datos convertidos a formato numérico.");
    }
}
