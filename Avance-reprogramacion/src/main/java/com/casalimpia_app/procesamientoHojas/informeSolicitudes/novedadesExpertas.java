/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class novedadesExpertas {
    public static void resaltarNovedad(Workbook wb) throws IOException {
        
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");

        // Crear un estilo de celda con relleno amarillo
        CellStyle yellowStyle = wb.createCellStyle();
        yellowStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Iterar sobre las filas (empezando desde la fila 1 para saltar el encabezado)
        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) {
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                // Obtener la celda de la columna M (índice 12)
                Cell cellN = row.getCell(12); // Columna M = índice 12

                if (cellN != null && cellN.getCellType() == Cell.CELL_TYPE_STRING) {
                    String valorNovedad = cellN.getStringCellValue();

                    // Si el valor es "Si", resaltar las celdas en las columnas J (índice 9) y K (índice 10)
                    if (valorNovedad.equalsIgnoreCase("Si")) {
                        Cell cellJ = row.getCell(9); // Columna J = índice 9
                        Cell cellK = row.getCell(10); // Columna K = índice 10

                        if (cellJ != null) {
                            cellJ.setCellStyle(yellowStyle); // Aplicar el estilo amarillo a la columna J
                        }

                        if (cellK != null) {
                            cellK.setCellStyle(yellowStyle); // Aplicar el estilo amarillo a la columna K
                        }
                    }
                }
            }
        }
        int EliminarColumnaN = 12; // Índice de la columna (empezando desde 0)
        for (Row fila : ws) {
            if (fila != null && fila.getCell(EliminarColumnaN) != null) {
                fila.removeCell(fila.getCell(EliminarColumnaN));
            }
        }
        System.out.println("Proceso completado. Las celdas de las columnas J y K han sido resaltadas donde corresponda.");
    }
}
