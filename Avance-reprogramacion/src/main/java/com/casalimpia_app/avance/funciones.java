/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.casalimpia_app.avance;

import static com.casalimpia_app.procesamientoHojas.hoja1.buscarVNombres.BUSCARVNombres;
import static com.casalimpia_app.procesamientoHojas.hoja1.buscarVNombresCedulas.BUSCARVNombresCedulas;
import static com.casalimpia_app.procesamientoHojas.hoja1.concatenarNombresApellidos.concatenacion;
import static com.casalimpia_app.procesamientoHojas.hoja1.intFormat.convertirTextoANumeroHoja1;
import static com.casalimpia_app.procesamientoHojas.hoja1.noServiceCopyPasteNewSheet.copiarFilasNoServiceNewSheet;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.IDCategorias.hogarOficina;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.dateFormat.formatearFechas;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.filtrarCiudades.filtrarCiudades;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.intFormat.convertirTextoANumero;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.novedadesExpertas.resaltarNovedad;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.duplicadosClientes.resaltarClientesDuplicados;
import static com.casalimpia_app.procesamientoHojas.informeSolicitudes.deleteExpertasSinOtro.eraseRows;
import static com.casalimpia_app.procesamientoHojas.maquillaje.headersStyles.estilosEncabezados;
import static com.casalimpia_app.procesamientoHojas.maquillaje.orderHoja1.reorganizeExcelHoja1;
import static com.casalimpia_app.procesamientoHojas.maquillaje.orderInformeSolicitudes.reorganizeExcelInformeSolicitudes;
import static com.casalimpia_app.procesamientoHojas.maquillaje.widthColumnSize.ajustarAnchoColumnas;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.filtrarExpertasNovedades.copyPasteNovedades;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.writeNews.escribirNovedades;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.writeNewsWithService.escribirNovedadesConServicio;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.cuidadoras.escribirCuidadoras;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.expertasEstandar.escribirExpertasEstandar;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.backup8And4.clasificarBackups;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.noServiceCopyPasteNews.copiarFilasNoServiceNovedades;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.orderExpertasSinServicio.reorganizeExcelExpertasSinServicio;
import static com.casalimpia_app.procesamientoHojas.expertasSinServicio.expertasNovedades.filtrarExpertasNovedadesAmarillo;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author jcavilaa
 */
public class funciones {
    public static void ejecucionFunciones(Workbook wb) throws IOException, Exception{
        //COLOCAR FUNCIÓN AQUÍ 
        filtrarExpertasNovedadesAmarillo(wb);
        resaltarNovedad(wb);
        filtrarCiudades(wb);
        convertirTextoANumero(wb);
        formatearFechas(wb);
        concatenacion(wb);
        convertirTextoANumeroHoja1(wb);
        BUSCARVNombresCedulas(wb);
        BUSCARVNombres(wb);
        //copiarFilasNoService(wb);
        estilosEncabezados(wb);
        ajustarAnchoColumnas(wb);
        reorganizeExcelInformeSolicitudes(wb);
        reorganizeExcelHoja1(wb);
        hogarOficina(wb);
        copiarFilasNoServiceNewSheet(wb);
        copyPasteNovedades(wb);
        escribirNovedades(wb);
        escribirNovedadesConServicio(wb);
        escribirCuidadoras(wb);
        escribirExpertasEstandar(wb);
        clasificarBackups(wb);
        reorganizeExcelExpertasSinServicio(wb);
        copiarFilasNoServiceNovedades(wb);
        resaltarClientesDuplicados(wb);
        eraseRows(wb);
    }
}
