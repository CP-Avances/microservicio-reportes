package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteAuditoria.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAuditoriaService {

    public byte[] generarReportePDF(ReporteAuditoriaRequest request) {

        // Constantes DRY
        final float[] WIDTHS_CABECERA = { 8f, 2f };
        final float[] WIDTHS_TABLA = { 1f, 3f, 2f, 2.5f, 3f, 2f, 2f, 2f, 6f, 6f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (respetando helpers)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("AUDITORÍA"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebra = ReporteUtil.colorZebraClaro();

            List<AuditoriaDTO> lista = request.getAuditorias();

            // Si no hay registros
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
                    new Phrase("PLATAFORMA: " + lista.get(0).getPlataforma(), ReporteUtil.fuenteTexto()));
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

            // Tabla principal de auditoría
            PdfPTable tabla = new PdfPTable(10);
            tabla.setWidthPercentage(100);
            tabla.setSpacingBefore(5f);
            tabla.setWidths(WIDTHS_TABLA);

            // Encabezados
            String[] encabezados = {
                    "ITEM", "PLATAFORMA", "USUARIO", "IP", "NOMBRE TABLA",
                    "ACCIÓN", "FECHA", "HORA", "DATOS ORIGINALES", "DATOS NUEVOS"
            };
            for (String enc : encabezados) {
                tabla.addCell(ReporteUtil.crearCelda(enc, ReporteUtil.fuenteEncabezado(), colorPrincipal));
            }

            // Cuerpo de la tabla
            AtomicInteger index = new AtomicInteger(1);
            for (AuditoriaDTO a : lista) {
                Color fondo = (index.get() % 2 == 0) ? zebra : Color.WHITE;

                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(index.getAndIncrement()), ReporteUtil.fuenteTexto(),
                        fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getPlataforma(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getUser_name(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getIp_address(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getTable_name(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getAction(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getFecha_hora_format(), ReporteUtil.fuenteTexto(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(a.getSolo_hora(), ReporteUtil.fuenteTexto(), fondo));

                PdfPCell celdaOriginal = ReporteUtil.crearCelda(a.getOriginal_data(), ReporteUtil.fuenteTexto(), fondo);
                celdaOriginal.setNoWrap(false);
                celdaOriginal.setMinimumHeight(20f);
                tabla.addCell(celdaOriginal);

                PdfPCell celdaNuevo = ReporteUtil.crearCelda(a.getNew_data(), ReporteUtil.fuenteTexto(), fondo);
                celdaNuevo.setNoWrap(false);
                celdaNuevo.setMinimumHeight(20f);
                tabla.addCell(celdaNuevo);
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper lanza IAEx, dejamos que el controller maneje 400
            throw e;
        } catch (Exception e) {
            // 500 uniforme con excepción de dominio
            throw new ReportBuildException("No se pudo generar ReporteAuditoria.pdf", e);
        } finally {
            // 4) Ciclo de recursos garantizado
            if (document != null && document.isOpen()) {
                try {
                    document.close();
                } catch (Exception ignore) {
                }
            }
            if (writer != null) {
                try {
                    writer.close();
                } catch (Exception ignore) {
                }
            }
            if (baos != null) {
                try {
                    baos.close();
                } catch (Exception ignore) {
                }
            }
        }
    }

}
