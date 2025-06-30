package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Feriado.FeriadoDTO;
import com.casapazmino.microservicio_reportes.model.Feriado.ReporteFeriadosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteFeriadosService {

    public byte[] generarReporteFeriadosPDF(ReporteFeriadosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate());
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
                logo.scaleToFit(100, 50);
                logo.setAlignment(Image.ALIGN_LEFT);
                document.add(logo);
            }

            // Empresa
            Paragraph empresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-30f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            // Título
            Paragraph titulo = new Paragraph("LISTA DE FERIADOS", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 16));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = new Color(204, 209, 209); // #CCD1D1

            // Tabla
            PdfPTable tabla = new PdfPTable(4);
            tabla.setWidthPercentage(65);
            tabla.setWidths(new float[]{2, 6, 4, 4});
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(crearCelda("CÓDIGO", colorPrincipal));
            tabla.addCell(crearCelda("DESCRIPCIÓN", colorPrincipal));
            tabla.addCell(crearCelda("FECHA", colorPrincipal));
            tabla.addCell(crearCelda("RECUPERACIÓN", colorPrincipal));

            // Filas
            List<FeriadoDTO> lista = request.getFeriados();
            for (int i = 0; i < lista.size(); i++) {
                FeriadoDTO f = lista.get(i);
                Color bgColor = (i % 2 == 0) ? colorZebra : null;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(f.getId()), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(f.getDescripcion(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(f.getFecha(), ReporteUtil.fuenteTexto(), bgColor));
                tabla.addCell(ReporteUtil.crearCelda(f.getFechaRecuperacion(), ReporteUtil.fuenteTexto(), bgColor));
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
