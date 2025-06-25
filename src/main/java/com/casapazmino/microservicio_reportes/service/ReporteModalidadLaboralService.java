package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteModalidadLaboralRequest;
import com.casapazmino.microservicio_reportes.model.ModalidadLaboralDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.Base64;
import java.util.List;

@Service
public class ReporteModalidadLaboralService {

    public byte[] generarReportePDF(ReporteModalidadLaboralRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(request.getUsuario(), request.getFraseMarcaAgua(), request.getColorPrincipal()));
            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Empresa
            Paragraph empresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-30f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            Paragraph titulo = new Paragraph("MODALIDAD LABORAL", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Color principal
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Tabla de datos
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(65);
            tabla.setWidths(new float[]{1, 5});
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("ITEM", ReporteUtil.fuenteEncabezado(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("MODALIDAD LABORAL", ReporteUtil.fuenteEncabezado(), colorPrincipal));

            // Filas con efecto zebra
            boolean zebra = false;
            List<ModalidadLaboralDTO> lista = request.getModalidades();
            for (ModalidadLaboralDTO modalidad : lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(modalidad.getId()), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(modalidad.getDescripcion(), ReporteUtil.fuenteTexto(), fondo));
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
