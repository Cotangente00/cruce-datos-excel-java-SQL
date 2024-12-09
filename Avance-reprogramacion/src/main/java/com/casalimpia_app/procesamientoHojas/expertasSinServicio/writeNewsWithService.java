package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;

public class writeNewsWithService {
    public static void escribirNovedadesConServicio(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas calendario");
        Sheet ws2 = wb.getSheet("INFORME SOLICITUDES");

        Iterator<Row> rowIterator = ws.iterator();

        // Usar un Map para almacenar todos los valores asociados a cada número
        Map<String, List<String>> datosExpertasCalendario = new HashMap<>();
        DataFormatter formatter = new DataFormatter();
    
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell celdaA = row.getCell(0);
            String numeroBuscar = formatter.formatCellValue(celdaA);
            String motivo = row.getCell(5).getStringCellValue();
    
            // Agregar el motivo a la lista asociada al número
            datosExpertasCalendario.computeIfAbsent(numeroBuscar, k -> new ArrayList<>()).add(motivo);
        }

        System.out.println(datosExpertasCalendario);

        // Iterar sobre la segunda hoja
        rowIterator = ws2.iterator();
        rowIterator.next(); // Saltar el encabezado
    
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell celdaM = row.getCell(9);
            String numeroBuscar = formatter.formatCellValue(celdaM);
    
            // Si el número tiene motivos asociados, escribirlos en las celdas adyacentes
            if (datosExpertasCalendario.containsKey(numeroBuscar)) {
                List<String> motivos = datosExpertasCalendario.get(numeroBuscar);
                int startColumn = 14; // Columna O
                for (int i = 0; i < motivos.size(); i++) {
                    row.createCell(startColumn + i).setCellValue(motivos.get(i));
                }
            }
        }
    }
}
