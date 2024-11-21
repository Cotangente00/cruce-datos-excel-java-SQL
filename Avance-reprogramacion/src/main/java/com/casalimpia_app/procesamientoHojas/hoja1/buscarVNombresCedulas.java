/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.hoja1;

import java.util.Collections;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class buscarVNombresCedulas {
    public static void BUSCARVNombresCedulas(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");
        Sheet ws2 = wb.getSheet("Hoja1");

        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator1 = ws.iterator();
        rowIterator1.next(); // Saltar el encabezado
        Iterator<Row> rowIterator2 = ws2.iterator();

        // Crear conjuntos para almacenar los números de documento
        Set<String> numerosHoja1 = new HashSet<>();
        Set<Map<String, String>> datosHoja2 = new HashSet<>();

        // Llenar los conjuntos con los datos
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Cell cell = row.getCell(9);
            DataFormatter formatter = new DataFormatter();
            String numeroBuscar = formatter.formatCellValue(cell);
            numerosHoja1.add(numeroBuscar);
        }
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell cell = row.getCell(3);
            DataFormatter formatter = new DataFormatter();
            String numeroBuscar = formatter.formatCellValue(cell);
            String nombre = row.getCell(4).getStringCellValue();
            datosHoja2.add(Collections.singletonMap(numeroBuscar, nombre));
        }

        // Iterar sobre los números de la Hoja1 y buscar coincidencias
        rowIterator1 = ws.iterator();
        rowIterator1.next(); // Saltar el encabezado
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Cell cell = row.getCell(9);
            DataFormatter formatter = new DataFormatter();
            String numeroBuscar = formatter.formatCellValue(cell);
            for (Map<String, String> dato : datosHoja2) {
                if (dato.containsKey(numeroBuscar)) {
                    row.createCell(12).setCellValue(numeroBuscar);
                    row.createCell(13).setCellValue(dato.get(numeroBuscar));
                    break;
                }
            }
        }

        int[] columnas = {12};

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
            }
        }
        System.out.println("Números de documento y nombres completos agregados exitosamente.");
    }
}
