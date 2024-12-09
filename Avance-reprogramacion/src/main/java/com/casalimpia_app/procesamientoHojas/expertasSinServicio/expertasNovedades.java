package com.casalimpia_app.procesamientoHojas.expertasSinServicio;

import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

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

        Set<String> numerosNovedades = new HashSet<>();

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
    }
}
