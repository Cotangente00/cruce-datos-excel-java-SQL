/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class IDCategorias {
    public static void hogarOficina (Workbook wb){
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");
        
        for (int i = 1; i <= ws.getLastRowNum(); i++) {
            Row row = ws.getRow(i);

            // Si la fila es nula, detener el ciclo
            if (row == null) {
                break;
            }
            
            // Leer la celda de la columna C (Subtipo)
            Cell cellC = row.getCell(2);
            if (cellC == null || cellC.getCellType() == Cell.CELL_TYPE_BLANK) {
                // detener el ciclo si se encuentra una vacía 
                break;
            } else if ("1".equals(cellC.getStringCellValue())) {
                cellC.setCellValue("Hogar");
            } else {
                cellC.setCellValue("Oficina");
            }
            
            
        }
    }
}
