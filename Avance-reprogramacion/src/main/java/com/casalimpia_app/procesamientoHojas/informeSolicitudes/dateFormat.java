/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.informeSolicitudes;

import java.io.IOException;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class dateFormat {
    public static void formatearFechas(Workbook wb) throws IOException {

        Sheet ws = wb.getSheetAt(0);

        // Crear un estilo de celda para formato de fecha DD/MM/YYYY
        CellStyle dateCellStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));

        // Iterar sobre las filas para convertir y formatear las fechas en la columna D (índice 3)
        for (int rowIndex = 1; rowIndex <= ws.getLastRowNum(); rowIndex++) { // Inicia en 1 para saltar el encabezado
            Row row = ws.getRow(rowIndex);
            if (row != null) {
                Cell cellFecha = row.getCell(3); // Columna D es el índice 3 (0-indexed)

                if (cellFecha != null && cellFecha.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    double numericValue = cellFecha.getNumericCellValue();


                    // Verificar si el número puede ser interpretado como una fecha
                    if (DateUtil.isValidExcelDate(numericValue)) {
                        // Convertir el valor numérico a una fecha
                        Date date = DateUtil.getJavaDate(numericValue);

                        // Cambiar el tipo de celda a fecha y aplicar el formato
                        cellFecha.setCellValue(date);
                        cellFecha.setCellStyle(dateCellStyle);
                    }
                }
            }
        }
        System.out.println("Proceso completado. Fechas formateadas.");
    }
}
