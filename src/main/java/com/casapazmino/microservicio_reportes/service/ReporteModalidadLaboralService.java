package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ModalidadLaboralDTO;
import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ReporteModalidadLaboralRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteModalidadLaboralService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteModalidadLaboralRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            //TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                request.getUsuario(),
                request.getFraseMarcaAgua(),
                request.getColorPrincipal()
            ));
            document.open();

            //LOGO DE EMPRESA
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            //TITULO DE EMPRESA
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            
            //TITULO DE REPORTE
            document.add(ReporteUtil.crearTituloReporte("MODALIDAD LABORAL"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            //TABLA DE LOS DATOS DE MODALIDAD LABORAL
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(new float[]{1, 5});
            tabla.setSpacingBefore(10f);

            //ENCABEZADOS DE LA TABLA
            tabla.addCell(ReporteUtil.crearCelda("ITEM", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("MODALIDAD LABORAL", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            //DATOS DE MODALIDAD LABORAL
            boolean zebra = false;
            List<ModalidadLaboralDTO> lista = request.getModalidades();
            for (ModalidadLaboralDTO modalidad : lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(modalidad.getId()), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(modalidad.getDescripcion(), ReporteUtil.fuenteTablaData(), fondo));
                zebra = !zebra;
            }

            document.add(tabla);
            document.close();
            writer.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}