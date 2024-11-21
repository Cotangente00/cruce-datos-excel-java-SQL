/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.avance;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.UIManager;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class guadarArchivo {
    public static void guardarArchivoVentana(Workbook wb) throws IOException {
        try {
            // Configurar estilo del JFileChooser
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Seleccionar ruta para guardar el archivo");
            fileChooser.setSelectedFile(new File("INFORME_SOLICITUDES.xlsx"));

            int userSelection = fileChooser.showSaveDialog(null);

            if (userSelection == JFileChooser.APPROVE_OPTION) {
                File archivoSeleccionado = fileChooser.getSelectedFile();
                try (FileOutputStream outputStream = new FileOutputStream(archivoSeleccionado)) {
                    wb.write(outputStream);
                    System.out.println("Archivo guardado en: " + archivoSeleccionado.getAbsolutePath());
                }
            } else {
                System.out.println("Guardado cancelado por el usuario.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
