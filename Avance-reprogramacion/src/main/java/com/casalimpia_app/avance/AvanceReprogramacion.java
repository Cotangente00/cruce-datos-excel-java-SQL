/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.casalimpia_app.avance;
import static com.casalimpia_app.avance.consultas.querys;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;

/**
 *
 * @author jcavilaa
 */
public class AvanceReprogramacion {
    
    public static void conexion() throws IOException, FileNotFoundException, FileNotFoundException, FileNotFoundException, Exception {
        try {
            // Cargar el controlador JDBC para SQL Server (ajusta según tu versión)
            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

            // URL de conexión, usuario y contraseña
            String url = "jdbc:sqlserver://192.168.1.3;databaseName=CASALIMPIA";
            String user = "fenix";
            String password = "Beck5388100NI";
            // Método para ejecutar las consultas
            querys(url, user, password);
            
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        } 
    }
    
    public static void main(String[] Args) throws IOException, Exception{
        conexion();
    }
}
