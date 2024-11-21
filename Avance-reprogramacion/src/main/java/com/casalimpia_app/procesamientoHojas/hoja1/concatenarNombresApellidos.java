/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.hoja1;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class concatenarNombresApellidos {
    public static void concatenacion (Workbook wb){
        Sheet ws = wb.getSheet("Hoja1");
        
        // Concatenar nombres y apellidos
        int columnaE = 4, columnaF = 5;
        for (Row row : ws) {
            Cell celdaE = row.getCell(columnaE);
            Cell celdaF = row.getCell(columnaF);
            if (celdaE != null && celdaF != null) {
                String nombre = celdaE.getStringCellValue();
                String apellido = celdaF.getStringCellValue();
                String nombreCompleto = nombre + " " + apellido;
                celdaE.setCellValue(nombreCompleto);
            }   
        }
        
        // Eliminar la columna F (índice 5)
        int eliminarColumnaN = 5;  // Índice de la columna (empezando desde 0)
        for (Row row : ws) {
            if (row != null && row.getCell(eliminarColumnaN) != null) {
                row.removeCell(row.getCell(eliminarColumnaN));
            }
        }
        
        // Mover todas las celdas a la izquierda
        for (Row row : ws) {
            Cell celdaActual = row.getCell(5);
            Cell celdaSiguiente = row.getCell(5 + 1);

            if (celdaSiguiente != null) {
                if (celdaActual == null) {
                    celdaActual = row.createCell(5);
                }
                copiarCelda(celdaSiguiente, celdaActual);
            } else if (celdaActual != null) {
                row.removeCell(celdaActual);
            }
            
        }
        // Eliminar la columna F (índice 5)
        int eliminarColumnaG = 6; // Índice de la columna (empezando desde 0)
        for (Row row : ws) {
            if (row != null && row.getCell(eliminarColumnaG) != null) {
                row.removeCell(row.getCell(eliminarColumnaG));
            }
        }
        
    }
        
    // Función para copiar el contenido de una celda a otra sin usar setCellType
    private static void copiarCelda(Cell desde, Cell hacia) {
        switch (desde.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                hacia.setCellValue(desde.getStringCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(desde)) {
                    hacia.setCellValue(desde.getDateCellValue());
                } else {
                    hacia.setCellValue(desde.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                hacia.setCellValue(desde.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                hacia.setCellFormula(desde.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK:
                hacia.setCellType(null);
                break;
            case Cell.CELL_TYPE_ERROR:
                hacia.setCellErrorValue(desde.getErrorCellValue());
                break;
            default:
                break;
        }
    }
}
