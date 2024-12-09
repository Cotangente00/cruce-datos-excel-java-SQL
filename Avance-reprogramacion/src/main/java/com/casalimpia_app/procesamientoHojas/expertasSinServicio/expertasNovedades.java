package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class expertasNovedades {
    public static void filtrarExpertasNovedadesAmarillo(Workbook wb) throws Exception {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES"); 
        Sheet ws2 = wb.createSheet("ExpertasNovedades");

        Iterator<Row> rowIterator = ws.iterator();
        rowIterator.next(); // Saltar encabezado
        Iterator<Row> rowIterator2 = ws2.iterator();


        List<String> numerosNovedades = new ArrayList<>();

        // Crear conjuntos para almacenar los números de documento

        // Llenar los conjuntos con los datos
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell novedad = row.getCell(12);
            String yesOrNot = novedad.getStringCellValue();
            while (novedad.getCellType() != Cell.CELL_TYPE_BLANK){
                if (yesOrNot.equals("Si") || yesOrNot.equals("si") || yesOrNot.equals("sí") || yesOrNot.equals("Sí") || yesOrNot.equals("SÍ") || yesOrNot.equals("SI")) { 
                    Cell numeros = row.getCell(9);
                    DataFormatter formatter = new DataFormatter();
                    String numeroBuscar = formatter.formatCellValue(numeros);
                    numerosNovedades.add(numeroBuscar);
                }
                break;
            }
            
        }
        
        for (String numero : numerosNovedades){
            System.out.println(numero);
        }        
        
        int rowIndex = 0;
        for (String numero : numerosNovedades) {
            Row row = ws2.createRow(rowIndex++);
            row.createCell(0).setCellValue(numero);
        }
        
        int[] columnas = {0};

        for (int i = 0; i <= ws2.getLastRowNum(); i++) { // Inicia en 1 para saltar el encabezado
            Row row = ws2.getRow(i);
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

    }
}
