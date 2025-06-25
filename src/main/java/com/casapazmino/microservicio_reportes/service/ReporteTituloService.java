package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.TituloDTO;
import com.casapazmino.microservicio_reportes.model.ReporteTitulosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteTituloService {

    public byte[] generarReporteTitulosPDF(ReporteTitulosRequest request) {
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
            Paragraph titulo = new Paragraph("LISTA DE TÍTULOS PROFESIONALES", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Tabla
            PdfPTable tabla = new PdfPTable(3);
            tabla.setWidthPercentage(90);
            tabla.setWidths(new float[]{2, 4, 6});
            tabla.setSpacingBefore(10f);

            // Encabezado
            tabla.addCell(crearCelda("CÓDIGO", colorPrincipal));
            tabla.addCell(crearCelda("NIVEL", colorPrincipal));
            tabla.addCell(crearCelda("NOMBRE", colorPrincipal));

            // Filas
            List<TituloDTO> lista = request.getTitulos();
            for (int i = 0; i < lista.size(); i++) {
                TituloDTO t = lista.get(i);
                Color bgColor = (i % 2 == 0) ? colorZebra : null;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(t.getId()), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(t.getNivel(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(t.getNombre(), ReporteUtil.fuenteTexto(), bgColor));
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
