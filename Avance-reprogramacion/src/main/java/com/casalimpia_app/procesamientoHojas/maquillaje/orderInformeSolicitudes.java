/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.maquillaje;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class orderInformeSolicitudes {
    public static void reorganizeExcelInformeSolicitudes(Workbook wb) throws IOException {
        Sheet originalSheet = wb.getSheetAt(0); // Obtener la primera hoja
        Sheet newSheet = wb.createSheet("ReorganizedSheet"); // Crear una nueva hoja para los datos reorganizados

        List<Row> rowsWithRepeatedJK = new ArrayList<>();
        List<Row> rowsWithEmptyMN = new ArrayList<>();
        List<Row> rowsWithData = new ArrayList<>();
        List<Row> data = new ArrayList<>();

        Set<String> uniqueJKValues = new HashSet<>();
        Set<String> repeatedJKValues = new HashSet<>();

        // Detectar filas con valores repetidos en las columnas J y K
        int rowIndex = 1; // Empezar desde la fila 2 (índice 1)
        while (true) {
            Row row = originalSheet.getRow(rowIndex);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK) {
                break; // Detener cuando se encuentre la primera celda vacía en la columna A
            }

            Cell cellJ = row.getCell(9);  // Columna J (índice 9)
            Cell cellK = row.getCell(10); // Columna K (índice 10)

            String valueJ = cellJ != null ? cellJ.toString() : "";
            String valueK = cellK != null ? cellK.toString() : "";

            String combinedValue = valueJ + "||" + valueK; // Combinar valores para identificar repeticiones

            if (!valueJ.isEmpty() && !valueK.isEmpty()) {
                if (!uniqueJKValues.add(combinedValue)) {
                    repeatedJKValues.add(combinedValue); // Registrar como repetido
                }
            }

            rowIndex++;
        }

        // Separar filas repetidas en J y K
        rowIndex = 1;
        while (true) {
            Row row = originalSheet.getRow(rowIndex);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK) {
                break;
            }

            Cell cellJ = row.getCell(9);
            Cell cellK = row.getCell(10);

            String valueJ = cellJ != null ? cellJ.toString() : "";
            String valueK = cellK != null ? cellK.toString() : "";
            String combinedValue = valueJ + "||" + valueK;

            if (repeatedJKValues.contains(combinedValue)) {
                rowsWithRepeatedJK.add(row); // Filas con datos repetidos en J y K
            } else {
                // Verificar columnas M y N
                Cell cellM = row.getCell(12);
                Cell cellN = row.getCell(13);

                boolean isCellMEmpty = (cellM == null || cellM.getCellType() == Cell.CELL_TYPE_BLANK);
                boolean isCellNEmpty = (cellN == null || cellN.getCellType() == Cell.CELL_TYPE_BLANK);

                if (isCellMEmpty && isCellNEmpty) {
                    rowsWithEmptyMN.add(row);
                } else {
                    rowsWithData.add(row);
                }
            }

            rowIndex++;
        }

        // Copiar filas con datos repetidos en J y K primero
        int newRowIndex = 1; // Comenzar desde la fila 2 en la nueva hoja
        for (Row row : rowsWithRepeatedJK) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        // Copiar filas con datos después
        for (Row row : rowsWithData) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        // Copiar filas con celdas vacías en M y N al final
        for (Row row : rowsWithEmptyMN) {
            copyRow(row, newSheet.createRow(newRowIndex++), wb);
        }

        // Eliminar datos de la hoja original, omitiendo los encabezados
        for (int rowIndex2 = 1; rowIndex2 <= originalSheet.getLastRowNum(); rowIndex2++) {
            Row row = originalSheet.getRow(rowIndex2);
            if (row != null) {
                originalSheet.removeRow(row);
            } else {
                break;
            }
        }

        // Almacenar todos los datos de la columna
        int rowIndex2 = 1; // Comenzar desde la fila 2 en la hoja nueva
        while (true) {
            Row row = newSheet.getRow(rowIndex2);
            if (row == null || row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK) {
                break; // Detener cuando se encuentre la primera celda vacía en la columna A
            }
            data.add(row);
            rowIndex2++;
        }

        int newRowIndex2 = 1; // Comenzar desde la fila 2 en la nueva hoja
        for (Row row : data) {
            copyRow(row, originalSheet.createRow(newRowIndex2++), wb);
        }

        wb.removeSheetAt(4);
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
