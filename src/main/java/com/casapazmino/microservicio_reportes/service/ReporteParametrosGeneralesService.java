package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Parametro.ParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ParametroDetalleDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosGeneralesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteParametrosGeneralesService {

    public byte[] generarReportePDF(ReporteParametrosGeneralesRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4);
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getNombreUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            // Logo
            try {
                Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                if (logo != null) {
                    logo.setAbsolutePosition(40, 770);
                    logo.scaleToFit(100, 60);
                    document.add(logo);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }

            // Título empresa
            Paragraph empresa = new Paragraph(request.getNombreEmpresa(), ReporteUtil.fuenteEncabezado());
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(10f);
            empresa.setSpacingAfter(2f);
            document.add(empresa);

            // Título módulo
            Paragraph titulo = new Paragraph("PARÁMETROS GENERALES", ReporteUtil.fuenteEncabezado());
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores y fuentes
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Font fuenteTexto = ReporteUtil.fuenteTexto();
            Font fuenteEncabezado = ReporteUtil.fuenteEncabezado();

            for (ParametroDTO parametro : request.getParametros()) {
                // Tabla cabecera del parámetro
                PdfPTable cabecera = new PdfPTable(2);
                cabecera.setWidthPercentage(100);
                cabecera.setSpacingBefore(10f);
                cabecera.setWidths(new float[] { 1f, 1f });

                PdfPCell celda1 = ReporteUtil.celdaIzquierda("PARÁMETRO: " + parametro.getDescripcion(), fuenteTexto);
                celda1.setBackgroundColor(colorPrincipal);
                celda1.setBorder(Rectangle.TOP | Rectangle.LEFT | Rectangle.BOTTOM); // SIN borde derecho

                PdfPCell celda2 = ReporteUtil.celdaDerecha("CÓDIGO PARÁMETRO: " + parametro.getId(), fuenteTexto);
                celda2.setBackgroundColor(colorPrincipal);
                celda2.setBorder(Rectangle.TOP | Rectangle.RIGHT | Rectangle.BOTTOM); // SIN borde izquierdo

                cabecera.addCell(celda1);
                cabecera.addCell(celda2);
                document.add(cabecera);

                // Tabla de detalles
                List<ParametroDetalleDTO> detalles = parametro.getDetalles();
                if (detalles != null && !detalles.isEmpty()) {
                    PdfPTable tabla = new PdfPTable(3);
                    tabla.setWidthPercentage(100);
                    tabla.setSpacingBefore(0f);
                    tabla.setSpacingAfter(2f);
                    tabla.setWidths(new float[] { 25, 30, 45 });

                    // Encabezados
                    tabla.addCell(ReporteUtil.crearCelda("CÓDIGO DETALLE", fuenteEncabezado, colorSecundario));
                    tabla.addCell(ReporteUtil.crearCelda("DETALLE", fuenteEncabezado, colorSecundario));
                    tabla.addCell(ReporteUtil.crearCelda("DESCRIPCIÓN", fuenteEncabezado, colorSecundario));

                    for (ParametroDetalleDTO d : detalles) {
                        // Todas las filas con fondo colorSecundario
                        tabla.addCell(ReporteUtil.celdaCentro(String.valueOf(d.getId()), fuenteTexto));
                        tabla.addCell(ReporteUtil.celdaCentro(d.getDescripcion(), fuenteTexto));
                        tabla.addCell(ReporteUtil.celdaCentro(d.getObservacion(), fuenteTexto));
                    }

                    document.add(tabla);
                }
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
