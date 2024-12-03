package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import static com.casalimpia_app.procesamientoHojas.maquillaje.widthColumnSize.ajustarColumnasManualmente;

import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class noServiceCopyPasteNews {
    public static void copiarFilasNoServiceNovedades(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");
        Sheet ws1 = wb.getSheet("Expertas Sin Servicio");

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

        
        
        // Recorrer las filas de la hoja "Hoja1" desde la fila 5 en adelante
        for (int i = 0; i <= ws1.getLastRowNum(); i++) {
            Row filaHoja1 = ws1.getRow(i);

            if (filaHoja1 != null) {


                Row filaINFORME_SOLICITUDES = ws.getRow(filaDestino);
                if (filaINFORME_SOLICITUDES == null) {
                    filaINFORME_SOLICITUDES = ws.createRow(filaDestino);
                }

                // Copiar A, B, C, D y E a L, M, N y O
                Cell celdaA = filaHoja1.getCell(0);  
                Cell celdaB = filaHoja1.getCell(1);  
                Cell celdaC = filaHoja1.getCell(2);  
                Cell celdaD = filaHoja1.getCell(3);  
                Cell celdaE = filaHoja1.getCell(4);
                Cell celdaF = filaHoja1.getCell(5);  
                Cell celdaG = filaHoja1.getCell(6);  
                Cell celdaH = filaHoja1.getCell(7);

                // Crear y asignar valores a las celdas M, N y O en "INFORME SOLICITUDES"
                Cell celdaL = filaINFORME_SOLICITUDES.createCell(11);  
                Cell celdaM = filaINFORME_SOLICITUDES.createCell(12);  
                System.out.println("valor de celda P después de copiar " + celdaM.toString());
                Cell celdaN = filaINFORME_SOLICITUDES.createCell(13);  
                Cell celdaO = filaINFORME_SOLICITUDES.createCell(14);  
                Cell celdaP = filaINFORME_SOLICITUDES.createCell(15);  
                Cell celdaQ = filaINFORME_SOLICITUDES.createCell(16);
                Cell celdaR = filaINFORME_SOLICITUDES.createCell(17);  
                Cell celdaS = filaINFORME_SOLICITUDES.createCell(18);

                System.out.println("Copiando fila: " + (i + 1) + " | A:" + (celdaA != null ? celdaA.toString() : "") + " | B:" + (celdaB != null ? celdaB.toString() : "") + " | C:" + (celdaC != null ? celdaC.toString() : "")  + " | D:" + (celdaD != null ? celdaD.toString() : "")
                + " | E:" + (celdaE != null ? celdaE.toString() : ""));

                // Asignar los valores a L, M, N y O si las celdas C, D, E y F no están vacías
                if (celdaA != null)  {
                    // Si es numérico, usar formato decimal para evitar notación científica
                    if (celdaA.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        celdaL.setCellValue(df.format(celdaA.getNumericCellValue()));
                    } else {
                        celdaL.setCellValue(celdaA.toString());
                    }
                }
                if (celdaB != null) {
                    // Si es numérico, usar formato decimal para evitar notación científica
                    if (celdaB.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        celdaM.setCellValue(df.format(celdaB.getNumericCellValue()));
                    } else {
                        celdaM.setCellValue(celdaB.toString());
                    }
                }
                if (celdaC != null) celdaN.setCellValue(celdaC.toString());
                if (celdaD != null) celdaO.setCellValue(celdaD.toString());
                if (celdaE != null) celdaP.setCellValue(celdaE.toString());
                if (celdaF != null) celdaQ.setCellValue(celdaF.toString());
                if (celdaG != null) celdaR.setCellValue(celdaG.toString());
                if (celdaH != null) celdaS.setCellValue(celdaH.toString());
                
                filaDestino++;  // Mover a la siguiente fila en Hoja1
            } else {
                System.err.println("No hay filas para copiar");
                break;
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

        ajustarColumnasManualmente(ws, 14);
        ajustarColumnasManualmente(ws, 15);
        ajustarColumnasManualmente(ws, 16);
        ajustarColumnasManualmente(ws, 17);
        ajustarColumnasManualmente(ws, 18); 

        wb.removeSheetAt(2);
        wb.removeSheetAt(2);
        
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
