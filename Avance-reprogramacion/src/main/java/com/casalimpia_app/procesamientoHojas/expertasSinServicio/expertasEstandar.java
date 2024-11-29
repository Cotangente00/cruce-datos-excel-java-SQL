package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.*;

import org.apache.poi.ss.usermodel.*;

public class expertasEstandar {
    public static void escribirExpertasEstandar (Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas Sin Servicio"); 

        Iterator<Row> rowIterator = ws.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell celdaE = row.getCell(4);

            if (celdaE == null || celdaE.getCellType() == Cell.CELL_TYPE_BLANK || celdaE.getStringCellValue() == "") {
                celdaE = row.createCell(4);
                celdaE.setCellValue("estandar");
            }
        }    
    }
}
