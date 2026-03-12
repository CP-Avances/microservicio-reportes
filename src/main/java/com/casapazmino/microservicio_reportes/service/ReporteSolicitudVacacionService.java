package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.SolicitudVacacion.AprobacionDTO;
import com.casapazmino.microservicio_reportes.model.SolicitudVacacion.ReporteSolicitudVacacionRequest;
import com.casapazmino.microservicio_reportes.model.SolicitudVacacion.SolicitudVacacionDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;

import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.Phrase;
import org.openpdf.text.Rectangle;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.List;

@Service
public class ReporteSolicitudVacacionService {

    public byte[] generarReporteSolicitudVacacionPDF(ReporteSolicitudVacacionRequest request) {
        final float[] WIDTHS_INFO_GENERAL = { 2.0f, 4.5f, 2.0f, 3.5f };
        final float[] WIDTHS_MOTIVO = { 2.0f, 4.5f, 2.0f, 3.5f };
        final float[] WIDTHS_FIRMAS = { 1f, 1f };

        final String TITULO_REPORTE = "SOLICITUD DE VACACIONES";

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            if (request == null) throw new IllegalArgumentException("request nulo");
            if (request.getSolicitud() == null) throw new IllegalArgumentException("solicitud nula");

            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO_REPORTE));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());

            final SolicitudVacacionDTO s = request.getSolicitud();

            // INFORMACIÓN GENERAL
            document.add(crearTituloSeccion("INFORMACIÓN GENERAL", colorPrincipal));

            PdfPTable info = new PdfPTable(4);
            info.setWidthPercentage(100);
            info.setWidths(WIDTHS_INFO_GENERAL);

            // FECHA y CÉDULA
            // Para vacaciones no siempre tienes "fecha_creacion"; usamos fecha_registro si viene
            addCeldaLabel(info, "FECHA:", colorPrincipal);
            addCeldaValor(info, fmtEcDateTime(s.getFecha_registro()));

            addCeldaLabel(info, "CÉDULA:", colorPrincipal);
            addCeldaValor(info, safe(s.getIdentificacion()));

            // CIUDAD y CÓDIGO
            addCeldaLabel(info, "CIUDAD:", colorPrincipal);
            addCeldaValor(info, safe(s.getNom_ciudad()));

            addCeldaLabel(info, "CÓDIGO:", colorPrincipal);
            addCeldaValor(info, safe(s.getCodigo()));

            // APELLIDOS
            addCeldaLabel(info, "APELLIDOS:", colorPrincipal);
            addCeldaValorColspan(info, safe(s.getApellido_emple()), 3);

            // NOMBRES
            addCeldaLabel(info, "NOMBRES:", colorPrincipal);
            addCeldaValorColspan(info, safe(s.getNombre_emple()), 3);

            // RÉGIMEN LABORAL
            addCeldaLabel(info, "RÉGIMEN LABORAL:", colorPrincipal);
            addCeldaValorColspan(info, safe(s.getNom_regimen()), 3);

            // DEPARTAMENTO y SUCURSAL
            addCeldaLabel(info, "DEPARTAMENTO:", colorPrincipal);
            addCeldaValor(info, safe(s.getNom_departamento()));

            addCeldaLabel(info, "SUCURSAL:", colorPrincipal);
            addCeldaValor(info, safe(s.getNom_sucursal()));

            info.setSpacingAfter(10f);
            document.add(info);

            // MOTIVO (en vacaciones se interpreta como rango de vacaciones)
            document.add(crearTituloSeccion("MOTIVO", colorPrincipal));

            PdfPTable motivo = new PdfPTable(4);
            motivo.setWidthPercentage(100);
            motivo.setWidths(WIDTHS_MOTIVO);

            // TIPO SOLICITUD y CÓDIGO en la misma fila
            addCeldaLabel(motivo, "TIPO SOLICITUD:", colorPrincipal);
            addCeldaValor(motivo, "VACACIONES");

            addCeldaLabel(motivo, "CÓDIGO SOLICITUD:", colorPrincipal);
            addCeldaValor(motivo, s.getId() == null ? "" : String.valueOf(s.getId()));

            // FECHA INICIO / FECHA TERMINACIÓN
            addCeldaLabel(motivo, "FECHA INICIO:", colorPrincipal);
            addCeldaValor(motivo, fmtEcDate(s.getFecha_inicio()));

            addCeldaLabel(motivo, "FECHA FIN:", colorPrincipal);
            addCeldaValor(motivo, fmtEcDate(s.getFecha_final()));

            // OBSERVACIÓN
            addCeldaLabel(motivo, "OBSERVACIÓN:", colorPrincipal);
            addCeldaValor(motivo, "");

            addCeldaLabel(motivo, "NÚMERO DE DÍAS:", colorPrincipal);
            addCeldaValor(motivo, s.getNumero_dias_totales() == null ? "" : String.valueOf(s.getNumero_dias_totales()));

            // OBSERVACIÓN AUTORIZACIÓN
            String obsAut = obtenerObservacionAutorizacion(request.getAprobaciones());
            addCeldaLabel(motivo, "OBSERVACIÓN\nAUTORIZACIÓN:", colorPrincipal);
            addCeldaValorColspan(motivo, obsAut, 3);

            motivo.setSpacingAfter(18f);
            document.add(motivo);

            // FIRMAS
            PdfPTable firmas = new PdfPTable(2);
            firmas.setWidthPercentage(75);
            firmas.setHorizontalAlignment(Element.ALIGN_CENTER);
            firmas.setWidths(WIDTHS_FIRMAS);

            String nombreSolicitante = (safe(s.getApellido_emple()) + " " + safe(s.getNombre_emple())).trim();
            PdfPCell solicitado = crearCajaFirma(
                    "Solicitado en el Sistema FullTime por:",
                    nombreSolicitante,
                    safe(s.getCargo())
            );

            List<AprobacionDTO> pasos = request.getAprobaciones();

            if (pasos == null || pasos.isEmpty()) {
                firmas.addCell(solicitado);

                PdfPCell empty = new PdfPCell(new Phrase(""));
                empty.setBorder(Rectangle.NO_BORDER);
                empty.setPaddingTop(6f);
                empty.setPaddingBottom(10f);
                empty.setPaddingLeft(6f);
                empty.setPaddingRight(6f);
                firmas.addCell(empty);
            } else {
                firmas.addCell(solicitado);
                pasos.sort(Comparator.comparingInt(a -> a.getOrden_paso() == null ? 0 : a.getOrden_paso()));
                for (AprobacionDTO a : pasos) {
                    if (a == null) continue;
                    firmas.addCell(crearCajaFirmaPaso(a));
                }

                if ((firmas.getRows().size() % 2) != 0) {
                    PdfPCell empty = new PdfPCell(new Phrase(""));
                    empty.setBorder(Rectangle.NO_BORDER);
                    empty.setPaddingTop(6f);
                    empty.setPaddingBottom(10f);
                    empty.setPaddingLeft(6f);
                    empty.setPaddingRight(6f);
                    firmas.addCell(empty);
                }
            }

            document.add(firmas);

            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar SolicitudVacacion.pdf", e);
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

    // ===================== HELPERS (copiados igual) =====================

    private PdfPTable crearTituloSeccion(String titulo, Color colorPrincipal) {
        PdfPTable t = new PdfPTable(1);
        t.setWidthPercentage(100);

        PdfPCell c = new PdfPCell(new Phrase(titulo, ReporteUtil.fuenteEncabezadoTablaData()));
        c.setBackgroundColor(colorPrincipal);
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setPaddingTop(4f);
        c.setPaddingBottom(5f);
        c.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.TOP | Rectangle.BOTTOM);

        t.addCell(c);
        t.setSpacingAfter(0f);
        t.setSpacingBefore(8f);
        return t;
    }

    private void addCeldaLabel(PdfPTable table, String texto, Color colorPrincipal) {
        PdfPCell c = new PdfPCell(new Phrase(texto, ReporteUtil.fuenteEncabezadoTablaData()));
        c.setBackgroundColor(Color.WHITE);
        c.setPaddingTop(4f);
        c.setPaddingBottom(4f);
        c.setPaddingLeft(5f);
        c.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.TOP | Rectangle.BOTTOM);
        table.addCell(c);
    }

    private void addCeldaValor(PdfPTable table, String texto) {
        PdfPCell c = new PdfPCell(new Phrase(safe(texto), ReporteUtil.fuenteEncabezado()));
        c.setBackgroundColor(Color.WHITE);
        c.setPaddingTop(4f);
        c.setPaddingBottom(4f);
        c.setPaddingLeft(5f);
        c.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.TOP | Rectangle.BOTTOM);
        table.addCell(c);
    }

    private void addCeldaValorColspan(PdfPTable table, String texto, int colspan) {
        PdfPCell c = new PdfPCell(new Phrase(safe(texto), ReporteUtil.fuenteEncabezado()));
        c.setBackgroundColor(Color.WHITE);
        c.setPaddingTop(4f);
        c.setPaddingBottom(4f);
        c.setPaddingLeft(5f);
        c.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.TOP | Rectangle.BOTTOM);
        c.setColspan(colspan);
        table.addCell(c);
    }

    private PdfPCell crearCajaFirma(String titulo, String nombre, String cargo) {
        PdfPTable box = new PdfPTable(1);
        box.setWidthPercentage(100);

        PdfPCell h = new PdfPCell(new Phrase(titulo, ReporteUtil.fuenteEncabezado()));
        h.setHorizontalAlignment(Element.ALIGN_CENTER);
        h.setPaddingTop(4f);
        h.setPaddingBottom(4f);
        h.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.TOP);

        PdfPCell cuerpo = new PdfPCell(new Phrase("\n\n\n", ReporteUtil.fuenteEncabezadoTablaData()));
        cuerpo.setBorder(Rectangle.LEFT | Rectangle.RIGHT);
        cuerpo.setMinimumHeight(80f);

        PdfPCell pieNombre = new PdfPCell(new Phrase(nombre, ReporteUtil.fuenteEncabezadoTablaData()));
        pieNombre.setHorizontalAlignment(Element.ALIGN_CENTER);
        pieNombre.setPaddingTop(4f);
        pieNombre.setPaddingBottom(2f);
        pieNombre.setBorder(Rectangle.LEFT | Rectangle.RIGHT);

        PdfPCell pieCargo = new PdfPCell(new Phrase(cargo, ReporteUtil.fuenteEncabezadoTablaData()));
        pieCargo.setHorizontalAlignment(Element.ALIGN_CENTER);
        pieCargo.setPaddingTop(0f);
        pieCargo.setPaddingBottom(6f);
        pieCargo.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.BOTTOM);

        box.addCell(h);
        box.addCell(cuerpo);
        box.addCell(pieNombre);
        box.addCell(pieCargo);

        PdfPCell wrapper = new PdfPCell(box);
        wrapper.setBorder(Rectangle.NO_BORDER);
        wrapper.setPaddingLeft(6f);
        wrapper.setPaddingRight(6f);
        wrapper.setPaddingTop(6f);
        wrapper.setPaddingBottom(10f);

        return wrapper;
    }

    private PdfPCell crearCajaFirmaConFecha(String titulo, String nombre, String cargo, String fechaHora) {
        PdfPTable box = new PdfPTable(1);
        box.setWidthPercentage(100);

        PdfPCell h = new PdfPCell(new Phrase(titulo, ReporteUtil.fuenteEncabezado()));
        h.setHorizontalAlignment(Element.ALIGN_CENTER);
        h.setPaddingTop(4f);
        h.setPaddingBottom(4f);
        h.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.TOP);

        PdfPCell cuerpo = new PdfPCell(new Phrase("\n\n\n", ReporteUtil.fuenteEncabezadoTablaData()));
        cuerpo.setBorder(Rectangle.LEFT | Rectangle.RIGHT);
        cuerpo.setMinimumHeight(80f);

        PdfPCell pieNombre = new PdfPCell(new Phrase(nombre, ReporteUtil.fuenteEncabezadoTablaData()));
        pieNombre.setHorizontalAlignment(Element.ALIGN_CENTER);
        pieNombre.setPaddingTop(4f);
        pieNombre.setPaddingBottom(2f);
        pieNombre.setBorder(Rectangle.LEFT | Rectangle.RIGHT);

        PdfPCell pieCargo = new PdfPCell(new Phrase(cargo, ReporteUtil.fuenteEncabezadoTablaData()));
        pieCargo.setHorizontalAlignment(Element.ALIGN_CENTER);
        pieCargo.setPaddingTop(0f);
        pieCargo.setPaddingBottom(2f);
        pieCargo.setBorder(Rectangle.LEFT | Rectangle.RIGHT);

        PdfPCell pieFecha = new PdfPCell(new Phrase(fechaHora, ReporteUtil.fuenteEncabezadoTablaData()));
        pieFecha.setHorizontalAlignment(Element.ALIGN_CENTER);
        pieFecha.setPaddingTop(0f);
        pieFecha.setPaddingBottom(6f);
        pieFecha.setBorder(Rectangle.LEFT | Rectangle.RIGHT | Rectangle.BOTTOM);

        box.addCell(h);
        box.addCell(cuerpo);
        box.addCell(pieNombre);
        box.addCell(pieCargo);
        box.addCell(pieFecha);

        PdfPCell wrapper = new PdfPCell(box);
        wrapper.setBorder(Rectangle.NO_BORDER);
        wrapper.setPaddingLeft(6f);
        wrapper.setPaddingRight(6f);
        wrapper.setPaddingTop(6f);
        wrapper.setPaddingBottom(10f);

        return wrapper;
    }

    private String tituloFirmaPorAccion(String accion) {
        if (accion == null) return "Validado en el Sistema FullTime por:";
        switch (accion) {
            case "PREAUTORIZA": return "Preautorizado en el Sistema FullTime por:";
            case "AUTORIZA":    return "Autorizado en el Sistema FullTime por:";
            case "RECHAZA":     return "Rechazado en el Sistema FullTime por:";
            default:            return "Validado en el Sistema FullTime por:";
        }
    }

    private PdfPCell crearCajaFirmaPaso(AprobacionDTO a) {
        String titulo = tituloFirmaPorAccion(a.getAccion());
        String nombre = safe(a.getEmpleado_nombre());
        String cargo = safe(a.getCargo_en_momento());
        String fecha = fmtEcDateTime(a.getFecha_hora_accion());
        return crearCajaFirmaConFecha(titulo, nombre, cargo, fecha);
    }

    private String safe(String s) {
        return (s == null || s.trim().isEmpty()) ? "" : s.trim();
    }

    private String fmtEcDateTime(String iso) {
        if (iso == null || iso.trim().isEmpty()) return "N/D";
        try {
            OffsetDateTime odt = OffsetDateTime.parse(iso);
            return odt.atZoneSameInstant(ZoneId.of("America/Guayaquil"))
                    .format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm:ss"));
        } catch (Exception e) {
            return "N/D";
        }
    }

    private String fmtEcDate(String iso) {
        if (iso == null || iso.trim().isEmpty()) return "N/D";
        try {
            OffsetDateTime odt = OffsetDateTime.parse(iso);
            return odt.atZoneSameInstant(ZoneId.of("America/Guayaquil"))
                    .format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (Exception e) {
            return "N/D";
        }
    }

    private String obtenerObservacionAutorizacion(List<AprobacionDTO> aprobaciones) {
        if (aprobaciones == null || aprobaciones.isEmpty()) return "";
        return aprobaciones.stream()
                .sorted(Comparator.comparingInt(a -> a.getOrden_paso() == null ? 0 : a.getOrden_paso()))
                .map(a -> a.getObservacion() == null ? "" : a.getObservacion().trim())
                .filter(x -> !x.isEmpty())
                .reduce((first, second) -> second)
                .orElse("");
    }
}