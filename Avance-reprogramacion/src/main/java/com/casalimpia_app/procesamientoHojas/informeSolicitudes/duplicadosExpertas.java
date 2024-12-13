package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class duplicadosExpertas {
    public static void resaltarExpertasDuplicadas(Workbook wb) throws IOException {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES");

        Iterator<Row> rowIterator = ws.iterator();
        rowIterator.next(); // Saltar encabezado

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell celdaJ = row.getCell(9); // Columna J (índice 9)
            Cell celdaK = row.getCell(10); // Columna K (índice 10)

            if (row == null) {
                break;
            } 

            
        }
    }
}
