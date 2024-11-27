    /*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.avance;

import static com.casalimpia_app.avance.funciones.ejecucionFunciones;
import static com.casalimpia_app.avance.guadarArchivo.guardarArchivoVentana;
import com.toedter.calendar.JDateChooser;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.JOptionPane; 

/**
 *
 * @author jcavilaa
 */
public class consultas {
    public static void querys(String url, String user, String password) throws SQLException, Exception{
        Workbook wb;
        
        // Abrir un selector de fecha para el usuario
        JDateChooser dateChooser = new JDateChooser();
        JOptionPane.showMessageDialog(null, dateChooser, "Seleccione una fecha", JOptionPane.PLAIN_MESSAGE);
        java.util.Date selectedDate = dateChooser.getDate();

        // Verificar que el usuario haya seleccionado una fecha
        if (selectedDate == null) {
            JOptionPane.showMessageDialog(null, "No se seleccionó ninguna fecha", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Formatear la fecha seleccionada al formato SQL (YYYY-MM-DD)
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        String fechaSeleccionada = dateFormat.format(selectedDate);
        
        try ( // Obtener una conexión
                Connection connection = DriverManager.getConnection(url, user, password)) {
                System.out.println("Conexión establecida");
                try ( // Crear un statement para ejecutar consultas
                        Statement statement = connection.createStatement()) {
                    String consultaSQL = "SELECT * FROM [CASALIMPIA].[pymesHogar].[visorSolicitudes] " +
                                     "WHERE fecha = '" + fechaSeleccionada + "' " +
                                     "AND ciudad IN ('Bogotá', 'Cota', 'Chia', 'Cajica', 'Chía', 'Cajicá', 'bogota')";
                                     
                    try (PreparedStatement preparedStatement = connection.prepareStatement(consultaSQL)) {
                    //preparedStatement.setString(1, fechaSeleccionada);

                    try (ResultSet resultSet = preparedStatement.executeQuery()) {
                        if (!resultSet.isBeforeFirst()) { // Verifica si hay resultados
                            // Si no hay registros, muestra un mensaje al usuario
                            JOptionPane.showMessageDialog(null,
                                "No hay datos para la fecha seleccionada: " + fechaSeleccionada,
                                "Sin resultados",
                                JOptionPane.INFORMATION_MESSAGE);

                            // Cancelar el proceso y regresar a la ventana inicial
                            return; // Detiene el proceso actual
                        } else {
                            wb = new XSSFWorkbook();
                            Sheet ws = wb.createSheet("INFORME SOLICITUDES");
                            // Crear los encabezados
                            Row headerRow = ws.createRow(0);
                            headerRow.createCell(0).setCellValue("Solicitud");
                            headerRow.createCell(1).setCellValue("Ref. Externa");
                            headerRow.createCell(2).setCellValue("Subtipo");
                            headerRow.createCell(3).setCellValue("Fechas");
                            headerRow.createCell(4).setCellValue("Cliente");
                            headerRow.createCell(5).setCellValue("Cliente Email");
                            headerRow.createCell(6).setCellValue("Tiempo");
                            headerRow.createCell(7).setCellValue("Horario");
                            headerRow.createCell(8).setCellValue("Dirección");
                            headerRow.createCell(9).setCellValue("Cedula Profesional");
                            headerRow.createCell(10).setCellValue("Profesional");
                            headerRow.createCell(11).setCellValue("Estado");
                            // Si hay registros, continúa procesando los datos
                            int rowNum = 1;
                            while (resultSet.next()) {
                            // Iterar sobre los resultados e imprimirlos
                            // Escribir los datos de la primera consulta
                          
                                Row row = ws.createRow(rowNum++);
                                row.createCell(0).setCellValue(resultSet.getString("id_transaccion"));
                                row.createCell(1).setCellValue(resultSet.getString("ref_externa"));
                                row.createCell(2).setCellValue(resultSet.getString("id_categoria"));
                                row.createCell(3).setCellValue(resultSet.getString("fecha"));
                                row.createCell(4).setCellValue(resultSet.getString("fullname"));
                                row.createCell(5).setCellValue(resultSet.getString("email"));
                                row.createCell(6).setCellValue(resultSet.getString("horas") + (" horas"));
                                row.createCell(7).setCellValue(resultSet.getString("horario"));
                                row.createCell(8).setCellValue(resultSet.getString("direccion"));
                                row.createCell(9).setCellValue(resultSet.getString("cedula"));
                                row.createCell(10).setCellValue(resultSet.getString("nombre"));
                                row.createCell(11).setCellValue(resultSet.getString("estado"));
                                row.createCell(12).setCellValue(resultSet.getString("tiene_Novedad"));
                                row.createCell(13).setCellValue(resultSet.getString("ciudad"));

                                /*
                                // Obtener los valores de las columnas y mostrarlos
                                //int id = resultSet.getInt("id");
                                String idTransaccion = resultSet.getString("id_transaccion");
                                String idRefExterna = resultSet.getString("ref_externa");
                                String tipo = resultSet.getString("tipo");
                                String fecha = resultSet.getString("fecha");
                                String cliente = resultSet.getString("fullname");
                                String clienteEmail = resultSet.getString("email");
                                String horas = resultSet.getString("horas");
                                String horario = resultSet.getString("horario");
                                String direccion = resultSet.getString("direccion");
                                String docProfesional = resultSet.getString("cedula");
                                String nombreProfesional = resultSet.getString("nombre");
                                String estado = resultSet.getString("estado");
                                String ciudad = resultSet.getString("ciudad");

                                System.out.println("ID transacción: " + idTransaccion + " Ref. Externa: " + idRefExterna + " Subtipo: " + tipo + " Fechas: " + fecha + " Cliente: " + cliente + " Cliente email: " + clienteEmail + " Horas: " + horas + " horas" + " Horario: " + horario + " Dirección :" + direccion + " Cedula profesional: " + docProfesional + " Nombre profesional: " + nombreProfesional + " Estado: " + estado + " Ciudad: " + ciudad);
                                */
                            }   // Cerrar el primer ResultSet
                        }
                        
                    }
                }
                    
                        
                System.out.println("//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////");
                // Escribir los datos de la segunda consulta
                try (ResultSet visorSupernumerarios = statement.executeQuery("SELECT * FROM [CASALIMPIA].[pymesHogar].[visorReporteSupernumerarios] vs " +
                                                                            "WHERE Coord = 'TCVA' " +
                                                                            "AND (" +
                                                                            "    DATENAME(WEEKDAY, '" + fechaSeleccionada + "') <> 'Saturday' " + // Reemplaza GETDATE() con la fecha seleccionada
                                                                            "    OR (DATENAME(WEEKDAY, '" + fechaSeleccionada + "') = 'Saturday' AND Horario <> '200') " +
                                                                            ");")) {
                    // Escribir los datos de la segunda consulta
                    Sheet ws1 = wb.createSheet("Hoja1");
                    int rowNum2 = 3;
                    while (visorSupernumerarios.next()) {
                        /*
                        // imprimir los resultados en la terminal
                        String docProfesional = visorSupernumerarios.getString("cedula");
                        String nombreProfesional = visorSupernumerarios.getString("nombre");
                        String apellidoProfesional = visorSupernumerarios.getString("apellido");
                        String coord = visorSupernumerarios.getString("especial");

                        //IMPRIMIR COLUMNAS
                        System.out.println("Número documento: " + docProfesional + " Nombres: " + nombreProfesional + " Apellidos: " + apellidoProfesional + " Coord: " + coord);
                        */

                        Row row = ws1.createRow(rowNum2++);
                        row.createCell(3).setCellValue(visorSupernumerarios.getString("cedula"));
                        row.createCell(4).setCellValue(visorSupernumerarios.getString("nombre"));
                        row.createCell(5).setCellValue(visorSupernumerarios.getString("apellido"));
                        row.createCell(6).setCellValue(visorSupernumerarios.getString("Especial"));
                    } // Cerrar el segundo ResultSet
                    // Cerrar el statement
                    // Cerrar la conexión
                }
                
                String consultaSQLCalendario = "SELECT * FROM [CASALIMPIA].[pymesHogar].[visorCalendarioExp] " +
                     "WHERE (CAST(SUBSTRING(FechaInicio, 1, 10) AS DATE) <= ? " +
                     "AND CAST(SUBSTRING(FechaFin, 1, 10) AS DATE) >= ?) " +
                     "OR (CAST(SUBSTRING(FechaInicio, 1, 10) AS DATE) = ? " +
                     "AND CAST(SUBSTRING(FechaFin, 1, 10) AS DATE) = ?)";

                try (PreparedStatement preparedStatement = connection.prepareStatement(consultaSQLCalendario)) {
                    // Asignar la fecha ingresada como parámetro
                    preparedStatement.setString(1, fechaSeleccionada); // Para fechaInicio <= fecha ingresada
                    preparedStatement.setString(2, fechaSeleccionada); // Para fechaFin >= fecha ingresada
                    preparedStatement.setString(3, fechaSeleccionada); // Para fechaInicio = fechaFin = fecha ingresada
                    preparedStatement.setString(4, fechaSeleccionada); // Para fechaInicio = fechaFin = fecha ingresada

                    try (ResultSet resultSet = preparedStatement.executeQuery()) {
                        if (!resultSet.isBeforeFirst()) {
                            JOptionPane.showMessageDialog(null,
                                "No hay datos para la fecha seleccionada: " + fechaSeleccionada,
                                "Sin resultados",
                                JOptionPane.INFORMATION_MESSAGE);
                            return;
                        } else {
                            Sheet wsExpertasCalendario = wb.createSheet("Expertas calendario");

                            // Crear los encabezados
                            Row headerRow = wsExpertasCalendario.createRow(0);
                            headerRow.createCell(0).setCellValue("Cedula");
                            // Agrega aquí los demás encabezados necesarios...

                            int rowNum = 1;
                            while (resultSet.next()) {
                                Row row = wsExpertasCalendario.createRow(rowNum++);
                                row.createCell(0).setCellValue(resultSet.getString("ID"));
                                // Agrega aquí los demás datos de las columnas necesarios...
                            }
                        }
                    }
                }
                
            }   
        }
        
        
        
        // Método para ejecutar las funciones globalmente
        ejecucionFunciones(wb);
        // Método para guardar el archivo nuevo Excel resultante con una ventana
        guardarArchivoVentana(wb);
        /*
        FileOutputStream outputStream;    
        outputStream = new FileOutputStream("C:/Users/JCAVILAA/Documents/NetBeansProjects/Avance-reprogramacion/src/main/java/com/casalimpia_app/avance/INFORME SOLICITUDES.xlsx");
        wb.write(outputStream);
        outputStream.close();
        */
    }
}

