package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import org.apache.poi.ss.usermodel.*;
import java.util.*;

public class cuidadoras {
    public static void escribirCuidadoras(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas Sin Servicio"); 

        
        for (int i = 0; i <= ws.getLastRowNum(); i++) {
            Row row = ws.getRow(i);

            // Si la fila es nula, detener el ciclo
            if (row == null) {
                break;
            }

            // Leer la celda de la columna D (Área)
            Cell celdaD = row.getCell(3);
            if (celdaD == null || celdaD.getCellType() == Cell.CELL_TYPE_BLANK) {
                // detener el ciclo si se encuentra una vacía
                break;
            }

            String area = celdaD.getStringCellValue();

            // Crear la celda en la columna correspondiente si no existe
            Cell cellResultado = row.getCell(4);
            
            if (cellResultado == null) {
               cellResultado = row.createCell(4);
            } else {
                continue;
            }
                
            if (area.equals("CUIDADORA")){
                cellResultado.setCellValue("CUIDADORA");
            }

        }   
    }
}
