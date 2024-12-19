package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class deleteExpertasSinOtro {
    public static void eraseRows(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");

        // Convertir columna O (índice 14)
        int[] columnas = {14};

        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                for (int colIndex : columnas) {
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
                break;
            }
        }
        
        Iterator<Row> rowIterator = ws.iterator();
        rowIterator.next();
        
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            
            if (row==null){
                break;
            }

            Cell columnaO = row.getCell(14); 
        
            if (columnaO == null || columnaO.getCellType() == Cell.CELL_TYPE_BLANK){
                break;
            } else if (columnaO.getCellType() == Cell.CELL_TYPE_NUMERIC){
                columnaO.setCellType(CellType.BLANK);
            }
        }
    }
}
