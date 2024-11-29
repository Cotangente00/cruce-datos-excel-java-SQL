package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

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

public class writeNews {
    public static void escribirNovedades(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas calendario"); 
        Sheet ws2 = wb.getSheet("Expertas Sin Servicio"); 

        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator2 = ws.iterator();
        rowIterator2.next(); // Saltar el encabezado

        Set<Map<String, String>> datosExpertasCalendario = new HashSet<>();

        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell celdaA = row.getCell(0);
            DataFormatter formatter = new DataFormatter();
            String numeroBuscar = formatter.formatCellValue(celdaA);
            String motivo = row.getCell(5).getStringCellValue();
            datosExpertasCalendario.add(Collections.singletonMap(numeroBuscar, motivo));
        }

        rowIterator2 = ws2.iterator();
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell celdaB = row.getCell(1);
            DataFormatter formatter = new DataFormatter();
            String numeroBuscar = formatter.formatCellValue(celdaB);
            for (Map<String, String> dato : datosExpertasCalendario) {
                if (dato.containsKey(numeroBuscar)) {
                    row.createCell(4).setCellValue(dato.get(numeroBuscar));
                    break;
                }
            }
        }
    }
}
