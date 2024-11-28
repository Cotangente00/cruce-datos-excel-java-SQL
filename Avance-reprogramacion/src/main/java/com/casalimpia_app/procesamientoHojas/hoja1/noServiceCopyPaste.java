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
public class noServiceCopyPaste {
    public static void copiarFilasNoService(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");
        Sheet ws1 = wb.getSheet("Hoja1");

        if (ws1 == null || ws == null) {
            System.out.println("Una de las hojas no existe.");
            return;
        } 

        // Formato para preservar números sin notación científica
        DecimalFormat df = new DecimalFormat("0");

        // Obtener la última fila de la tabla en Hoja1 (A1:L*)
        int ultimaFilawsINFORME_SOLICITUDES = obtenerUltimaFilaTabla(ws);

        // La primera fila donde comenzaremos a copiar en la columna M es 10 filas después de la última fila de la tabla
        int filaDestino = ultimaFilawsINFORME_SOLICITUDES + 10;

        int primeraFilaHoja1 = 3;
        
        // Recorrer las filas de la hoja "Hoja1" desde la fila 5 en adelante
        for (int i = primeraFilaHoja1; i <= ws1.getLastRowNum(); i++) {
            Row filaHoja1 = ws1.getRow(i);

            if (filaHoja1 != null) {
                Cell celdaH = filaHoja1.getCell(7);  // Columna H es el índice 7

                // Si la celda H está vacía, copiar los datos de D, E y F a P, Q y R en Hoja1
                if (celdaH == null || celdaH.getCellType() == Cell.CELL_TYPE_BLANK) {
                    Row filaINFORME_SOLICITUDES = ws.getRow(filaDestino);
                    if (filaINFORME_SOLICITUDES == null) {
                        filaINFORME_SOLICITUDES = ws.createRow(filaDestino);
                    }

                    // Copiar C, D, E y F a L, M, N y O
                    Cell celdaC = filaHoja1.getCell(2);  // Columna C es el índice 2
                    Cell celdaD = filaHoja1.getCell(3);  // Columna D es el índice 3
                    Cell celdaE = filaHoja1.getCell(4);  // Columna E es el índice 4
                    Cell celdaF = filaHoja1.getCell(5);  // Columna F es el índice 5

                    // Crear y asignar valores a las celdas M, N y O en "INFORME SOLICITUDES"
                    Cell celdaL = filaINFORME_SOLICITUDES.createCell(11);  // Columna L es el índice 11
                    Cell celdaM = filaINFORME_SOLICITUDES.createCell(12);  // Columna M es el índice 12
                    System.out.println("valor de celda P después de copiar " + celdaM.toString());
                    Cell celdaN = filaINFORME_SOLICITUDES.createCell(13);  // Columna N es el índice 13
                    Cell celdaO = filaINFORME_SOLICITUDES.createCell(14);  // Columna O es el índice 14

                    System.out.println("Copiando fila: " + (i + 1) + " | D:" + (celdaD != null ? celdaD.toString() : "") + " | E:" + (celdaE != null ? celdaE.toString() : "") + " | F:" + (celdaF != null ? celdaF.toString() : ""));

                    // Asignar los valores a L, M, N y O si las celdas C, D, E y F no están vacías
                    if (celdaC != null)  {
                        // Si es numérico, usar formato decimal para evitar notación científica
                        if (celdaC.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            celdaL.setCellValue(df.format(celdaC.getNumericCellValue()));
                        } else {
                            celdaL.setCellValue(celdaC.toString());
                        }
                    }
                    if (celdaD != null) {
                        // Si es numérico, usar formato decimal para evitar notación científica
                        if (celdaD.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            celdaM.setCellValue(df.format(celdaD.getNumericCellValue()));
                        } else {
                            celdaM.setCellValue(celdaD.toString());
                        }
                    }
                    if (celdaE != null) celdaN.setCellValue(celdaE.toString());
                    if (celdaF != null) celdaO.setCellValue(celdaF.toString());

                    filaDestino++;  // Mover a la siguiente fila en Hoja1
                }
            }
        }
        
        int[] columna = {11, 12};

        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = ws.getRow(rowIndex);
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
            } else {
                continue;
            }
        }
        
    }
    
    // Método para obtener la última fila de la tabla en la hoja "INFORME SOLICITUDES"
    public static int obtenerUltimaFilaTabla(Sheet wsINFORME_SOLICITUDES) {
        int ultimaFila = 0;

        for (int i = 0; i <= wsINFORME_SOLICITUDES.getLastRowNum(); i++) {
            Row fila = wsINFORME_SOLICITUDES.getRow(i);
            if (fila != null) {
                for (int j = 0; j <= 11; j++) {  // Revisar las columnas de A a L (índices 0 a 11)
                    Cell celda = fila.getCell(j);
                    if (celda != null && celda.getCellType() != Cell.CELL_TYPE_BLANK) {
                        ultimaFila = i;  // Actualizar la última fila no vacía
                        break;
                    }
                }
            }
        }

        return ultimaFila;
    }
}
