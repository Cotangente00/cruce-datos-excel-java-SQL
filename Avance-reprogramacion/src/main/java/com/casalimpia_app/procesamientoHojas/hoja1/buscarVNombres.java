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
public class buscarVNombres {
    public static void BUSCARVNombres(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Hoja1"); // Hoja1
        Sheet ws2 = wb.getSheet("INFORME SOLICITUDES"); // INFORME SOLICITUDES

        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator1 = ws.iterator();
        Iterator<Row> rowIterator2 = ws2.iterator();
        rowIterator2.next(); // Saltar el encabezado de INFORME SOLICITUDES

        // Crear conjuntos para almacenar los números de documento
        Set<String> numerosHoja1 = new HashSet<>();
        Set<Map<String, String>> datosINFORME_SOLICITUDES = new HashSet<>();

        // Llenar los conjuntos con los datos
        DataFormatter formatter = new DataFormatter();

        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Cell cell = row.getCell(3); // Columna de números en Hoja1
            String numeroBuscar = formatter.formatCellValue(cell);
            numerosHoja1.add(numeroBuscar);
        }

        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell cellNumero = row.getCell(9); // Columna de números en INFORME SOLICITUDES
            Cell cellNombre = row.getCell(10); // Columna de nombres en INFORME SOLICITUDES

            String numeroBuscar = formatter.formatCellValue(cellNumero);
            String nombre = "";

            // Verificar si la celda del nombre no es nula y asignar valor
            if (cellNombre != null) {
                if (cellNombre.getCellType() == Cell.CELL_TYPE_STRING) {
                    nombre = cellNombre.getStringCellValue();
                } else if (cellNombre.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    nombre = String.valueOf(cellNombre.getNumericCellValue());
                }
            }

            if (!numeroBuscar.isEmpty() && !nombre.isEmpty()) {
                datosINFORME_SOLICITUDES.add(Collections.singletonMap(numeroBuscar, nombre));
            }
        }

        // Iterar sobre los números de la Hoja1 y buscar coincidencias
        rowIterator1 = ws.iterator();
        while (rowIterator1.hasNext()) {
            Row row = rowIterator1.next();
            Cell cell = row.getCell(3); // Columna de números en Hoja1
            String numeroBuscar = formatter.formatCellValue(cell);
            for (Map<String, String> dato : datosINFORME_SOLICITUDES) {
                if (dato.containsKey(numeroBuscar)) {
                    row.createCell(7).setCellValue(dato.get(numeroBuscar)); // Colocar el nombre en la columna 7
                    break;
                }
            }
        }


        System.out.println("Nombres completos agregados exitosamente.");
    }
}
