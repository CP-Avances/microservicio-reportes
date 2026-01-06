package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteAuditoria.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
public class ReporteAuditoriaService {

    public byte[] generarReportePDF(ReporteAuditoriaRequest request) {

        final float[] WIDTHS_CABECERA = { 8f, 2f };
        final float[] WIDTHS_TABLA = { 1.3f, 3f, 2f, 2.3f, 3f, 2f, 2f, 2f, 6f, 6f };

        // ✅ Ajusta este número según tu data (300-800 recomendado)
        final int CHUNK_ROWS = 500;

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("AUDITORÍA"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebra = ReporteUtil.colorZebraClaro();

            List<AuditoriaDTO> lista = request.getAuditorias();

            if (lista == null || lista.isEmpty()) {
                document.add(new Paragraph("No hay registros para mostrar.", ReporteUtil.fuenteTexto()));
                document.close();
                return baos.toByteArray();
            }

            // Cabecera: PLATAFORMA + N° registros
            PdfPTable cabecera = new PdfPTable(2);
            cabecera.setWidthPercentage(100);
            cabecera.setWidths(WIDTHS_CABECERA);
            cabecera.setSpacingBefore(10f);

            PdfPCell celdaPlataforma = new PdfPCell(
                    new Phrase("PLATAFORMA: " + safe(lista.get(0).getPlataforma()), ReporteUtil.fuenteTexto()));
            celdaPlataforma.setBackgroundColor(colorSecundario);
            celdaPlataforma.setBorder(Rectangle.TOP | Rectangle.LEFT | Rectangle.BOTTOM);
            celdaPlataforma.setPadding(5f);
            cabecera.addCell(celdaPlataforma);

            PdfPCell celdaCantidad = new PdfPCell(
                    new Phrase("N° Registros: " + lista.size(), ReporteUtil.fuenteTexto()));
            celdaCantidad.setBackgroundColor(colorSecundario);
            celdaCantidad.setBorder(Rectangle.TOP | Rectangle.RIGHT | Rectangle.BOTTOM);
            celdaCantidad.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaCantidad.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaCantidad.setPadding(5f);
            cabecera.addCell(celdaCantidad);

            document.add(cabecera);

            // ===== ✅ Tabla por bloques =====
            PdfPTable tabla = crearTablaAuditoria(colorPrincipal, WIDTHS_TABLA, WIDTHS_TABLA.length);

            int index = 1;
            int rowsEnBloque = 0;

            for (AuditoriaDTO a : lista) {
                Color fondo = (index % 2 == 0) ? zebra : Color.WHITE;

                // celdas normales
                tabla.addCell(ReporteUtil.celdaDataCentro(String.valueOf(index), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getPlataforma()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getUser_name()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getIp_address()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getTable_name()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getAction()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getFecha_hora_format()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(safe(a.getSolo_hora()), fondo));

                // ✅ columnas grandes: limita tamaño para evitar explosión de memoria
                PdfPCell celdaOriginal = ReporteUtil.celdaDataCentro(truncar(safe(a.getOriginal_data()), 2500), fondo);
                celdaOriginal.setNoWrap(false);
                celdaOriginal.setMinimumHeight(20f);
                tabla.addCell(celdaOriginal);

                PdfPCell celdaNuevo = ReporteUtil.celdaDataCentro(truncar(safe(a.getNew_data()), 2500), fondo);
                celdaNuevo.setNoWrap(false);
                celdaNuevo.setMinimumHeight(20f);
                tabla.addCell(celdaNuevo);

                index++;
                rowsEnBloque++;

                // ✅ flush cada CHUNK_ROWS
                if (rowsEnBloque >= CHUNK_ROWS) {
                    document.add(tabla);
                    document.add(Chunk.NEWLINE);

                    // recrear tabla (nueva instancia) para liberar memoria del bloque anterior
                    tabla = crearTablaAuditoria(colorPrincipal, WIDTHS_TABLA, WIDTHS_TABLA.length);
                    rowsEnBloque = 0;
                }
            }

            // flush final si quedó algo
            if (rowsEnBloque > 0) {
                document.add(tabla);
            }

            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteAuditoria.pdf", e);
        } finally {
            if (document != null && document.isOpen()) {
                try { document.close(); } catch (Exception ignore) {}
            }
            if (writer != null) {
                try { writer.close(); } catch (Exception ignore) {}
            }
            if (baos != null) {
                try { baos.close(); } catch (Exception ignore) {}
            }
        }
    }

    private PdfPTable crearTablaAuditoria(Color colorPrincipal, float[] widths, int cols) throws Exception {
        PdfPTable tabla = new PdfPTable(cols);
        tabla.setWidthPercentage(100);
        tabla.setSpacingBefore(5f);
        tabla.setWidths(widths);

        // ✅ hace que el header se repita por página
        tabla.setHeaderRows(1);

        // ✅ ayuda a partir filas en páginas (cuando hay texto largo)
        tabla.setSplitLate(false);
        tabla.setSplitRows(true);
        tabla.setKeepTogether(false);

        String[] encabezados = {
                "ITEM", "PLATAFORMA", "USUARIO", "IP", "NOMBRE TABLA",
                "ACCIÓN", "FECHA", "HORA", "DATOS ORIGINALES", "DATOS NUEVOS"
        };
        for (String enc : encabezados) {
            tabla.addCell(ReporteUtil.crearCelda(enc, ReporteUtil.fuenteEncabezado(), colorPrincipal));
        }
        return tabla;
    }

    private static String safe(String s) {
        return s == null ? "" : s;
    }

    private static String truncar(String s, int max) {
        if (s == null) return "";
        if (s.length() <= max) return s;
        return s.substring(0, max) + " ...";
    }


}
