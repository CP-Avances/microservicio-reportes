package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ProvinciaDTO;
import com.casapazmino.microservicio_reportes.model.ReporteProvinciasRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteProvinciasService {

    public byte[] generarReportePDF(ReporteProvinciasRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Empresa
            Paragraph empresa = new Paragraph(request.getEmpresa(),
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-30f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            Paragraph titulo = new Paragraph("LISTA DE PROVINCIAS",
                    FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Tabla
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(60); // centrada
            tabla.setWidths(new float[]{2, 4});
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("PAÍS", ReporteUtil.fuenteEncabezado(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("PROVINCIAS", ReporteUtil.fuenteEncabezado(), colorPrincipal));

            // Filas
            boolean zebra = false;
            List<ProvinciaDTO> lista = request.getProvincias();
            for (ProvinciaDTO provincia : lista) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(provincia.getPais(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(provincia.getNombre(), ReporteUtil.fuenteTexto(), fondo));
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
