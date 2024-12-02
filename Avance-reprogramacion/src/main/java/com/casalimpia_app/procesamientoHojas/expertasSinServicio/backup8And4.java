package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.*;
import org.apache.poi.ss.usermodel.*;


public class backup8And4 {
    public static void clasificarBackups (Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("Expertas Sin Servicio");

        Iterator<Row> rowIterator = ws.iterator();

        while (rowIterator.hasNext()){
            Row row = rowIterator.next();
            Cell celdaA = row.getCell(0);
            Cell celdaE = row.getCell(4);
            Cell celdaF = row.getCell(5);
            Cell celdaG = row.getCell(6);
            Cell celdaH = row.getCell(7);
            //
            /*
            if (celdaE == null || celdaF == null || celdaG == null || celdaH == null || celdaE.getCellType() == Cell.CELL_TYPE_BLANK || celdaF.getCellType() == Cell.CELL_TYPE_BLANK || celdaG.getCellType() == Cell.CELL_TYPE_BLANK || celdaH.getCellType() == Cell.CELL_TYPE_BLANK){
                continue;
            } 
                 */

            if (celdaA.getNumericCellValue() == 120 && celdaE.getStringCellValue().contains("backup")) {
                celdaE.setCellValue(celdaE.getStringCellValue() + " - 4 horas");
            } else if (celdaA.getNumericCellValue() == 180 && celdaE.getStringCellValue().contains("backup")){
                celdaE.setCellValue(celdaE.getStringCellValue() + " - 6 horas");
            } else if (celdaE.getStringCellValue().contains("backup")){
                celdaE.setCellValue(celdaE.getStringCellValue() + " - 8 horas");
            }

            if (celdaF == null || celdaF.getCellType() == Cell.CELL_TYPE_BLANK || celdaF.getStringCellValue() == "") { 
                continue;
            } else if (celdaA.getNumericCellValue() == 120 && celdaF.getStringCellValue().contains("backup")) {
                celdaF.setCellValue(celdaF.getStringCellValue() + " - 4 horas");
            } else if (celdaA.getNumericCellValue() == 180 && celdaF.getStringCellValue().contains("backup")){
                celdaF.setCellValue(celdaF.getStringCellValue() + " - 6 horas");
            } else if (celdaF.getStringCellValue().contains("backup")){
                celdaF.setCellValue(celdaF.getStringCellValue() + " - 8 horas");
            }

            if (celdaG == null || celdaG.getCellType() == Cell.CELL_TYPE_BLANK || celdaG.getStringCellValue() == "") { 
                continue;
            } else if (celdaA.getNumericCellValue() == 120 && celdaG.getStringCellValue().contains("backup")) {
                celdaG.setCellValue(celdaG.getStringCellValue() + " - 4 horas");
            } else if (celdaA.getNumericCellValue() == 180 && celdaG.getStringCellValue().contains("backup")){
                celdaG.setCellValue(celdaG.getStringCellValue() + " - 6 horas");
            } else if (celdaG.getStringCellValue().contains("backup")){
                celdaG.setCellValue(celdaG.getStringCellValue() + " - 8 horas");
            }

            if (celdaH == null || celdaH.getCellType() == Cell.CELL_TYPE_BLANK || celdaH.getStringCellValue() == "") { 
                continue;
            } else if (celdaA.getNumericCellValue() == 120 && celdaH.getStringCellValue().contains("backup")) {
                celdaH.setCellValue(celdaH.getStringCellValue() + " - 4 horas");
            } else if (celdaA.getNumericCellValue() == 180 && celdaH.getStringCellValue().contains("backup")){
                celdaH.setCellValue(celdaH.getStringCellValue() + " - 6 horas");
            } else if (celdaH.getStringCellValue().contains("backup")){
                celdaH.setCellValue(celdaH.getStringCellValue() + " - 8 horas");
            }
        }
    }
}
