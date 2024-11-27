/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.maquillaje;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class orderHoja1 {
    public static void reorganizeExcelHoja1(Workbook wb) throws IOException {
        Sheet originalSheet = wb.getSheet("Hoja1");  // Obtener la primera hoja
        Sheet newSheet = wb.createSheet("ReorganizedSheetHoja1");  // Crear una nueva hoja para los datos reorganizados

        List<Row> rowsWithEmptyH = new ArrayList<>();
        List<Row> rowsWithData = new ArrayList<>();
        List<Row> data = new ArrayList<>();

        // Iterar sobre la columna D para determinar el rango de filas
        int rowIndex = 3;  // Empezar desde la fila 4 (índice 3)
        while (true) {
            Row row = originalSheet.getRow(rowIndex);
            if (row == null || row.getCell(3) == null || row.getCell(3).getCellType() == Cell.CELL_TYPE_BLANK) {
                break;  // Detener cuando se encuentre la primera celda vacía en la columna D
            }

            // Verificar las celdas en la columna H (índice 7 respectivamente)
            Cell cellH = row.getCell(7);
            
            boolean isCellHEmpty = (cellH == null || cellH.getCellType() == Cell.CELL_TYPE_BLANK);

            // Si ambas la celda H está vacías, agregar la fila a la lista correspondiente
            if (isCellHEmpty) {
                rowsWithEmptyH.add(row);
            } else {
                rowsWithData.add(row);
            }

            rowIndex++;
        }

        // Copiar filas con datos primero
        int newRowIndex = 3;  // Comenzar desde la fila 4 en la nueva hoja
        for (Row row : rowsWithData) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        // Copiar filas con celdas vacías en H al final
        for (Row row : rowsWithEmptyH) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        //eliminar datos de la hoja original omitiendo los encabezados
        for (int rowIndex2 = 3; rowIndex2 <= originalSheet.getLastRowNum(); rowIndex2++) {
            Row row = originalSheet.getRow(rowIndex2);
            if (row != null) {
                originalSheet.removeRow(row);
            } else {
                break;
            }
        }


        // Almacenar todos los datos de la columna
        int rowIndex2 = 3;  // Comenzar desde la fila 4 en la hoja nueva
        while (true) {
            Row row = newSheet.getRow(rowIndex2);
            if (row == null || row.getCell(3) == null || row.getCell(3).getCellType() == Cell.CELL_TYPE_BLANK) {
                break;  // Detener cuando se encuentre la primera celda vacía en la columna A
            }
            data.add(row);
            rowIndex2++;
        }   

        int newRowIndex2 = 3; // Comenzar desde la fila 2 en la nueva hoja
        for (Row row : data) {
            copyRow(row, originalSheet.createRow(newRowIndex2++), wb);
        }

        wb.removeSheetAt(3);
    }

    // Método para copiar el contenido de una fila a otra
    public static void copyRow(Row sourceRow, Row targetRow, Workbook wb) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell sourceCell = sourceRow.getCell(i);
            Cell targetCell = targetRow.createCell(i);

            if (sourceCell != null) {
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

                CellStyle newCellStyle = wb.createCellStyle();
                newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
                targetCell.setCellStyle(newCellStyle);
            }
        }
    }
}
