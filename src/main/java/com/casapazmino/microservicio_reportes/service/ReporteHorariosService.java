package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.DetalleHorarioDTO;
import com.casapazmino.microservicio_reportes.model.HorarioDTO;
import com.casapazmino.microservicio_reportes.model.ReporteHorariosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteHorariosService {

    public byte[] generarReportePDF(ReporteHorariosRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4, 40, 40, 50, 50);
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
                logo.scaleAbsolute(100, 50);
                logo.setAlignment(Image.LEFT);
                document.add(logo);
            }

            // Empresa y título
            Paragraph empresa = new Paragraph(request.getEmpresa(), FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14));
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingBefore(-30f);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            Paragraph titulo = new Paragraph("LISTA DE HORARIOS", FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12));
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = new Color(229, 231, 233); // #E5E7E9

            for (HorarioDTO h : request.getHorarios()) {
                // Tabla contenedora de todo el bloque
                PdfPTable bloque = new PdfPTable(1);
                bloque.setWidthPercentage(100);
                bloque.setSpacingBefore(10f);

                // Cabecera (tabla 3 columnas)
                PdfPTable cabecera = new PdfPTable(3);
                cabecera.setWidthPercentage(100);
                cabecera.setWidths(new float[]{5, 5, 5});

                cabecera.addCell(celdaHorario("HORARIO: " + h.getNombre(), colorPrincipal, Rectangle.TOP | Rectangle.LEFT));
                cabecera.addCell(celdaHorario("HORAS DE TRABAJO: " + h.getHoraTrabajo(), colorPrincipal, Rectangle.TOP));
                cabecera.addCell(celdaHorario("MINUTOS DE ALIMENTACIÓN: " + h.getMinutosComida(), colorPrincipal, Rectangle.TOP | Rectangle.RIGHT));

                cabecera.addCell(celdaHorario("CÓDIGO: " + h.getCodigo(), colorPrincipal, Rectangle.LEFT));
                cabecera.addCell(celdaHorario("HORARIO NOCTURNO: " + (h.isNoturno() ? "Sí" : "No"), colorPrincipal, Rectangle.NO_BORDER));
                cabecera.addCell(celdaHorario("DOCUMENTO: " + (h.getDocumento() != null ? h.getDocumento() : ""), colorPrincipal, Rectangle.RIGHT));

                PdfPCell celdaCabecera = new PdfPCell(cabecera);
                celdaCabecera.setPadding(0);
                celdaCabecera.setBorder(Rectangle.NO_BORDER);
                bloque.addCell(celdaCabecera);

                if (h.getDetalles() != null && !h.getDetalles().isEmpty()) {
                    // Título "DETALLES"
                    PdfPTable tituloDetalles = new PdfPTable(1);
                    tituloDetalles.setWidthPercentage(100);
                    PdfPCell celdaDetalles = new PdfPCell(new Phrase("DETALLES", ReporteUtil.fuenteEncabezado()));
                    celdaDetalles.setBackgroundColor(colorSecundario);
                    celdaDetalles.setHorizontalAlignment(Element.ALIGN_CENTER);
                    celdaDetalles.setPaddingTop(5f);
                    celdaDetalles.setPaddingBottom(5f);
                    celdaDetalles.setBorder(Rectangle.BOX);
                    tituloDetalles.addCell(celdaDetalles);

                    PdfPCell celdaTitulo = new PdfPCell(tituloDetalles);
                    celdaTitulo.setPadding(0);
                    celdaTitulo.setBorder(Rectangle.NO_BORDER);
                    bloque.addCell(celdaTitulo);

                    // Tabla de detalles
                    PdfPTable tabla = new PdfPTable(7);
                    tabla.setWidthPercentage(100);
                    tabla.setWidths(new float[]{1.5f, 2.5f, 2.5f, 5, 2, 2.5f, 2.5f});

                    // Encabezados
                    String[] headers = {"ORDEN", "HORA", "TOLERANCIA", "ACCIÓN", "OTRO DÍA", "MINUTOS ANTES", "MINUTOS DESPUÉS"};
                    for (String col : headers) {
                        tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezado(), colorSecundario));
                    }

                    // Filas con efecto zebra
                    boolean zebra = false;
                    for (DetalleHorarioDTO d : h.getDetalles()) {
                        Color fondo = zebra ? zebraColor : null;
                        zebra = !zebra;

                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getOrden()), ReporteUtil.fuenteTexto(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getHora(), ReporteUtil.fuenteTexto(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getTolerancia() != null ? d.getTolerancia().toString() : "", ReporteUtil.fuenteTexto(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getTipoAccionShow(), ReporteUtil.fuenteTexto(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.isSegundoDia() ? "Sí" : "No", ReporteUtil.fuenteTexto(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getMinutosAntes()), ReporteUtil.fuenteTexto(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getMinutosDespues()), ReporteUtil.fuenteTexto(), fondo));
                    }

                    PdfPCell celdaTabla = new PdfPCell(tabla);
                    celdaTabla.setPadding(0);
                    celdaTabla.setBorder(Rectangle.NO_BORDER);
                    bloque.addCell(celdaTabla);
                }

                document.add(bloque);
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private PdfPCell celdaHorario(String texto, Color fondo, int border) {
        Font fuente = FontFactory.getFont(FontFactory.HELVETICA, 9);
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(5f);
        celda.setBorder(border);
        return celda;
    }
}
