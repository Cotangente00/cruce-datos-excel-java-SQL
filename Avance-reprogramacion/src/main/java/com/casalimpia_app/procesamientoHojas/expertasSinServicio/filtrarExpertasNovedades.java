package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;


public class filtrarExpertasNovedades {
    public static void copyPasteNovedades(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas Calendario"); 
        Sheet ws2 = wb.getSheet("Expertas Sin Servicio"); 

        Iterator<Row> rowIterator1 = ws2.iterator();
        Iterator<Row> rowIterator2 = ws.iterator();
        rowIterator2.next(); // Saltar encabezado

        List<Double> numDocExpertasWithoutService = new ArrayList<>();

        // Llenar la lista con los números de documento 
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Double numeroAsistencia = row.getCell(1).getNumericCellValue(); 
            numDocExpertasWithoutService.add(numeroAsistencia);
        }
        /*
        for (double numero : numDocExpertasWithoutService){
            System.out.println(numero);
        }
         */
        
        // Crear una lista para almacenar las filas a eliminar
        List<Integer> rowsToDelete = new ArrayList<>();

        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell numerosDoc = row.getCell(0);

            // Verificar si la celda existe y si su tipo es numérico
            if (numerosDoc != null && numerosDoc.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                double numeroDoc = numerosDoc.getNumericCellValue();

                if (!numDocExpertasWithoutService.contains(numeroDoc)) {
                    rowsToDelete.add(row.getRowNum());
                }
            }
        }

        // Eliminar filas en orden inverso para evitar problemas con los índices
        for (int i = rowsToDelete.size() - 1; i >= 0; i--) {
            int rowIndex = rowsToDelete.get(i);
            Row rowToRemove = ws.getRow(rowIndex);
            if (rowToRemove != null) {
                ws.removeRow(rowToRemove);

                // Eliminar también la fila física si está dentro del rango de datos
                int lastRowNum = ws.getLastRowNum();
                if (rowIndex < lastRowNum) {
                    ws.shiftRows(rowIndex + 1, lastRowNum, -1);
                }
            }
        }       
    }
}
