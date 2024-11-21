/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.procesamientoHojas.maquillaje;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class headersStyles {
    public static void estilosEncabezados(Workbook wb) throws IOException {
        Sheet ws = wb.getSheet("INFORME SOLICITUDES"); //Hoja INFORME_SOLICITUDES
        Row fila = ws.getRow(0); //Fila A

        // Crear un estilo con negrita y subrayado
        CellStyle estilo = wb.createCellStyle();
        Font fuente = wb.createFont();
        fuente.setBold(true);
        fuente.setUnderline(Font.U_SINGLE);
        estilo.setFont(fuente);
        
        //Aplicar los estilos por cada celda de la fila
        for (Cell celda : fila) {
            celda.setCellStyle(estilo);
        }
    }
}
