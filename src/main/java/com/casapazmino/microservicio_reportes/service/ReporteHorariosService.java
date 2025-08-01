package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Horario.DetalleHorarioDTO;
import com.casapazmino.microservicio_reportes.model.Horario.HorarioDTO;
import com.casapazmino.microservicio_reportes.model.Horario.ReporteHorariosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteHorariosService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteHorariosRequest request) {
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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE HORARIOS"));

            //COLORES DE LA EMPRESA USADOS EN EL REPORTE
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            for (HorarioDTO h : request.getHorarios()) {
                // Tabla contenedora
                PdfPTable bloque = new PdfPTable(1);
                bloque.setWidthPercentage(100);
                bloque.setSpacingBefore(10f);

                //TABLA CABEZERA DE CADA HORARIO
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
                    PdfPCell celdaDetalles = new PdfPCell(new Phrase("DETALLES", ReporteUtil.fuenteEncabezadoTablaData()));
                    celdaDetalles.setBackgroundColor(colorSecundario);
                    celdaDetalles.setHorizontalAlignment(Element.ALIGN_CENTER);
                    celdaDetalles.setPadding(5);
                    celdaDetalles.setBorder(Rectangle.BOX);
                    tituloDetalles.addCell(celdaDetalles);

                    PdfPCell celdaTitulo = new PdfPCell(tituloDetalles);
                    celdaTitulo.setPadding(0);
                    celdaTitulo.setBorder(Rectangle.NO_BORDER);
                    bloque.addCell(celdaTitulo);

                    // Tabla de detalles
                    PdfPTable tabla = new PdfPTable(7);
                    tabla.setWidthPercentage(100);
                    tabla.setWidths(new float[]{1.5f, 1.5f, 2.5f, 4.4f, 2, 3, 3});

                    String[] headers = {"ORDEN", "HORA", "TOLERANCIA", "ACCIÓN", "OTRO DÍA", "MINUTOS ANTES", "MINUTOS DESPUÉS"};
                    for (String col : headers) {
                        tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                    }

                    boolean zebra = false;
                    for (DetalleHorarioDTO d : h.getDetalles()) {
                        Color fondo = zebra ? zebraColor : null;
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getOrden()), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getHora(), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getTolerancia() != null ? d.getTolerancia().toString() : "", ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getTipoAccionShow(), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.isSegundoDia() ? "Sí" : "No", ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getMinutosAntes()), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getMinutosDespues()), ReporteUtil.fuenteTablaData(), fondo));
                        zebra = !zebra;
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
