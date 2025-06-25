package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.DiscapacidadDTO;
import com.casapazmino.microservicio_reportes.model.ReporteDiscapacidadesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteDiscapacidadService {

    public byte[] generarReporteDiscapacidadesPDF(ReporteDiscapacidadesRequest request) {
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
            Paragraph titulo = new Paragraph("LISTA DE DISCAPACIDADES", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209);

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(70);
            tabla.setWidths(new float[]{2, 6});
            tabla.setSpacingBefore(10f);

            tabla.addCell(crearCelda("CÓDIGO", colorPrincipal));
            tabla.addCell(crearCelda("NOMBRE", colorPrincipal));

            List<DiscapacidadDTO> lista = request.getDiscapacidades();
            for (int i = 0; i < lista.size(); i++) {
                DiscapacidadDTO d = lista.get(i);
                Color bgColor = (i % 2 == 0) ? colorZebra : null;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(d.getNombre(), ReporteUtil.fuenteTexto(), bgColor));
            }

            document.add(tabla);
            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell crearCelda(String texto, Color bgColor) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, ReporteUtil.fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setBackgroundColor(bgColor);
        return celda;
    }
}
