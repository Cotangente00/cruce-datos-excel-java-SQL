package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.*;
import org.apache.poi.ss.usermodel.*;

public class writeNews {
    public static void escribirNovedades(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas calendario");
        Sheet ws2 = wb.getSheet("Expertas Sin Servicio");
    
        // Obtener las columnas de interés como iteradores
        Iterator<Row> rowIterator2 = ws.iterator();
        rowIterator2.next(); // Saltar el encabezado
    
        // Usar un Map para almacenar todos los valores asociados a cada número
        Map<String, List<String>> datosExpertasCalendario = new HashMap<>();
        DataFormatter formatter = new DataFormatter();
    
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell celdaA = row.getCell(0);
            String numeroBuscar = formatter.formatCellValue(celdaA);
            String motivo = row.getCell(5).getStringCellValue();
    
            // Agregar el motivo a la lista asociada al número
            datosExpertasCalendario.computeIfAbsent(numeroBuscar, k -> new ArrayList<>()).add(motivo);
        }
    
        System.out.println(datosExpertasCalendario);
    
        // Iterar sobre la segunda hoja
        rowIterator2 = ws2.iterator();
        rowIterator2.next(); // Saltar el encabezado
    
        while (rowIterator2.hasNext()) {
            Row row = rowIterator2.next();
            Cell celdaB = row.getCell(1);
            String numeroBuscar = formatter.formatCellValue(celdaB);
    
            // Si el número tiene motivos asociados, escribirlos en las celdas adyacentes
            if (datosExpertasCalendario.containsKey(numeroBuscar)) {
                List<String> motivos = datosExpertasCalendario.get(numeroBuscar);
                int startColumn = 4; // Columna E
                for (int i = 0; i < motivos.size(); i++) {
                    row.createCell(startColumn + i).setCellValue(motivos.get(i));
                }
            }
        }
    }
}
