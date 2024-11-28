/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.hoja1;

import java.text.DecimalFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class noServiceCopyPasteNewSheet {
    public static void copiarFilasNoServiceNewSheet(Workbook wb){
        Sheet ws = wb.createSheet("Expertas Sin Servicio");
        Sheet ws1 = wb.getSheet("Hoja1");
        int filaDestino = 0;
        
        // Formato para preservar números sin notación científica
        DecimalFormat df = new DecimalFormat("0");
        
        // Recorrer las filas de la hoja "Hoja1" desde la fila 5 en adelante
        for (int i = 0; i <= ws1.getLastRowNum(); i++) {
            Row filaHoja1 = ws1.getRow(i);

            if (filaHoja1 != null) {
                Cell celdaH = filaHoja1.getCell(7);  // Columna H es el índice 7

                // Si la celda H está vacía, copiar los datos de D, E y F a A, B y C en Hoja1
                if (celdaH == null || celdaH.getCellType() == Cell.CELL_TYPE_BLANK) {
                    Row filaExpertasWithoutService = ws.getRow(filaDestino);
                    if (filaExpertasWithoutService == null) {
                        filaExpertasWithoutService = ws.createRow(filaDestino);
                    }

                    // Copiar D, E y F a A, B y C
                    Cell celdaC = filaHoja1.getCell(2);  // Columna C es el índice 2
                    Cell celdaD = filaHoja1.getCell(3);  // Columna D es el índice 3
                    Cell celdaE = filaHoja1.getCell(4);  // Columna E es el índice 4
                    Cell celdaF = filaHoja1.getCell(5);  // Columna F es el índice 5
                    
                    // Crear y asignar valores a las celdas P, Q y R en "INFORME SOLICITUDES"
                    Cell celda1 = filaExpertasWithoutService.createCell(0);  // Columna A es el índice 0
                    Cell celda2 = filaExpertasWithoutService.createCell(1);  // Columna B es el índice 1
                    Cell celda3 = filaExpertasWithoutService.createCell(2);  // Columna C es el índice 2
                    Cell celda4 = filaExpertasWithoutService.createCell(3);  // Columna C es el índice 3
                    
                    System.out.println("Copiando fila: " + (i + 1) + " | D:" + (celdaD != null ? celdaD.toString() : "") + " | E:" + (celdaE != null ? celdaE.toString() : "") + " | F:" + (celdaF != null ? celdaF.toString() : ""));
                    if (celdaC != null) {
                        // Si es numérico, usar formato decimal para evitar notación científica
                        if (celdaC.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            celda1.setCellValue(df.format(celdaC.getNumericCellValue()));
                        } else {
                            celda1.setCellValue(celdaC.toString());
                        }
                    }
                    if (celdaD != null) {
                        // Si es numérico, usar formato decimal para evitar notación científica
                        if (celdaD.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            celda2.setCellValue(df.format(celdaD.getNumericCellValue()));
                        } else {
                            celda2.setCellValue(celdaD.toString());
                        }
                    }
                    if (celdaE != null) celda3.setCellValue(celdaE.toString());
                    if (celdaF != null) celda4.setCellValue(celdaF.toString());

                    filaDestino++;  // Mover a la siguiente fila en Hoja1
                }
            }
        }
        
        int[] columna = {0,1};

        for (int i = 0; i <= ws.getLastRowNum(); i++) { 
            Row row = ws.getRow(i);
            if (row != null) {
                for (int colIndex : columna) {
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
        
    }   
}
