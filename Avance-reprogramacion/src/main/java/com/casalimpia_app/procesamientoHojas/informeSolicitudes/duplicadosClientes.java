package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.*;


public class duplicadosClientes {
    public static void resaltarClientesDuplicados(Workbook wb) throws IOException {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");

        // Mapa para almacenar los valores de la columna E y sus filas correspondientes
        Map<String, List<Integer>> valoresYFilas = new HashMap<>();

        Iterator<Row> rowIterator = ws.iterator();
        rowIterator.next();
        int filaActual = 1; // Inicializar el número de fila
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell celdaE = row.getCell(4); // Columna E

            if (celdaE != null && celdaE.getCellType() == Cell.CELL_TYPE_STRING) {
                String valorCelda = celdaE.getStringCellValue();

                // Obtener o crear la lista de filas para el valor actual
                List<Integer> filas = valoresYFilas.getOrDefault(valorCelda, new ArrayList<>());
                filas.add(filaActual);
                valoresYFilas.put(valorCelda, filas);

                filaActual++;
            }
        }

        // Aplicar formato de resaltado a todas las celdas duplicadas
        for (Map.Entry<String, List<Integer>> entry : valoresYFilas.entrySet()) {
            List<Integer> filas = entry.getValue();
            if (filas.size() > 1) { // Si hay más de una fila, es un duplicado
                CellStyle estiloDuplicado = wb.createCellStyle();
                estiloDuplicado.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                estiloDuplicado.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                for (int fila : filas) {
                    Row row = ws.getRow(fila);
                    row.getCell(4).setCellStyle(estiloDuplicado);
                }
            }
        }
    }
}
