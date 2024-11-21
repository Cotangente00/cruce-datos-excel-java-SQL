/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class filtrarCiudades {
    public static void filtrarCiudades(Workbook wb) throws IOException {

        Sheet ws = wb.getSheet("INFORME SOLICITUDES");

        // Índice de la columna "Ciudad" (N es la columna 13, 0-indexed)
        int columnaCiudadIndex = 13;
        int columnaOIndex = 14;

        // Crear un estilo de celda para las filas cuyas ciudades son "Soacha" y vacías
        CellStyle style = wb.createCellStyle();
        Font font = wb.createFont();
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setUnderline(Font.U_SINGLE);
        style.setFont(font);

        // Iterar sobre las filas y eliminar las que no cumplan con el criterio
        for (int rowIndex = ws.getLastRowNum(); rowIndex >= 1; rowIndex--) {  // Empieza desde el final para evitar problemas con el desplazamiento de filas y saltando el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                Cell cellCiudad = row.getCell(columnaCiudadIndex);
                String valorCiudad = (cellCiudad != null) ? cellCiudad.getStringCellValue().trim() : "";
                if (valorCiudad.equalsIgnoreCase("soacha")) {
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }
                    cellColumnaO.setCellValue("Soacha(Validar Servicio)");
                    cellColumnaO.setCellStyle(style);
                } else if (valorCiudad.isEmpty() || valorCiudad.equalsIgnoreCase("") || valorCiudad == null) {
                    Cell cellColumnaO = row.getCell(columnaOIndex);
                    if (cellColumnaO == null) {
                        cellColumnaO = row.createCell(columnaOIndex);
                    }
                    cellColumnaO.setCellValue("Ciudad vacía(Confirmar)");
                    cellColumnaO.setCellStyle(style);
                } 
            }
        }

        // Eliminar la columna N (índice 13)
        int eliminarColumnaN = 13;  // Índice de la columna (empezando desde 0)
        for (Row fila : ws) {
            if (fila != null && fila.getCell(eliminarColumnaN) != null) {
                fila.removeCell(fila.getCell(eliminarColumnaN));
            }
        }

        //System.out.println("Proceso completado. Filas filtradas.");
        
        /*
        // Eliminar todas las filas cuyo valor en la columna A esté vacío
        for (int rowIndex2 = ws.getLastRowNum(); rowIndex2 >= 1; rowIndex2--) {
            Row row = ws.getRow(rowIndex2);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK) {
                ws.removeRow(row);
            }
        }
        */
    }
}
