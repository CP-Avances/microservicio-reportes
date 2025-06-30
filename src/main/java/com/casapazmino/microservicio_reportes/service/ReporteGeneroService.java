package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Genero.GeneroDTO;
import com.casapazmino.microservicio_reportes.model.Genero.ReporteGenerosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteGeneroService {

    public byte[] generarReporteGenerosPDF(ReporteGenerosRequest request) {
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
            Paragraph empresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-30f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            Paragraph titulo = new Paragraph("LISTA DE GÉNEROS", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Tabla
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(60);
            tabla.setSpacingBefore(10f);
            tabla.setWidths(new float[]{2, 4});
            tabla.setHorizontalAlignment(Element.ALIGN_CENTER);

            // Encabezado
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezado(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("GÉNERO", ReporteUtil.fuenteEncabezado(), colorPrincipal));

            // Filas
            List<GeneroDTO> generos = request.getGeneros();
            for (int i = 0; i < generos.size(); i++) {
                GeneroDTO genero = generos.get(i);
                Color fondo = (i % 2 == 0) ? colorZebra : Color.WHITE;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(genero.getId()), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(genero.getGenero(), ReporteUtil.fuenteTexto(), fondo));
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
