package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Horario.DetalleHorarioDTO;
import com.casapazmino.microservicio_reportes.model.Horario.HorarioDTO;
import com.casapazmino.microservicio_reportes.model.Horario.ReporteHorariosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;


@Service
public class ReporteHorariosService {

    //METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteHorariosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS_CABECERA = { 5f, 5f, 5f };
        final float[] WIDTHS_DETALLES = { 1.5f, 1.5f, 2.5f, 4.4f, 2f, 3f, 3f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                request.getUsuario(),
                request.getFraseMarcaAgua(),
                request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes)
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE HORARIOS"));

            // Colores
            Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor      = ReporteUtil.colorZebraClaro();

            // Iteración de horarios
            for (HorarioDTO h : request.getHorarios()) {

                // Bloque contenedor
                PdfPTable bloque = new PdfPTable(1);
                bloque.setWidthPercentage(100);
                bloque.setSpacingBefore(10f);

                // --- CABECERA HORARIO ---
                PdfPTable cabecera = new PdfPTable(3);
                cabecera.setWidthPercentage(100);
                cabecera.setWidths(WIDTHS_CABECERA);

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

                // --- DETALLES DEL HORARIO ---
                if (h.getDetalles() != null && !h.getDetalles().isEmpty()) {

                    // Título de sección
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

                    // Tabla detalles
                    PdfPTable tabla = new PdfPTable(7);
                    tabla.setWidthPercentage(100);
                    tabla.setWidths(WIDTHS_DETALLES);

                    // Encabezados
                    String[] headers = {"ORDEN", "HORA", "TOLERANCIA", "ACCIÓN", "OTRO DÍA", "MINUTOS ANTES", "MINUTOS DESPUÉS"};
                    for (String col : headers) {
                        tabla.addCell(ReporteUtil.crearCelda(col, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                    }

                    // Cuerpo
                    boolean zebra = false;
                    for (DetalleHorarioDTO d : h.getDetalles()) {
                        Color fondo = zebra ? zebraColor : Color.WHITE;
                        tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getOrden()), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(d.getHora(), ReporteUtil.fuenteTablaData(), fondo));
                        tabla.addCell(ReporteUtil.crearCelda(
                            d.getTolerancia() != null ? d.getTolerancia().toString() : "",
                            ReporteUtil.fuenteTablaData(), fondo));
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

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // entrada inválida → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteHorarios.pdf", e);
        } finally {
            // 4) Ciclo de recursos garantizado
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

    // =========================
    //           XLSX
    // =========================
    public byte[] generarReporteXLSX(ReporteHorariosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Horarios"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;

        // Merges B1:N1 ... B5:N5 => (row 0..4, col 1..13)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 13;

        final String TITULO_REPORTE = "LISTA DE HORARIOS";

        final String[] HEADERS = {
            "ITEM","HORARIO","CÓDIGO","HORAS DE TRABAJO","MINUTOS DE ALIMENTACIÓN",
            "HORARIO NOTURNO","DOCUMENTO","ORDEN","HORA","TOLERANCIA",
            "ACCIÓN","OTRO DÍA","MINUTOS ANTES","MINUTOS DESPUÉS"
        };
        final int[] ANCHOS = { 10,20,20,20,20,20,20,20,20,20,20,20,30,30 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener encabezado visible

            // 1) Logo A1:B5 (si existe)
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges B1:N1 ... B5:N5
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, TITULO_REPORTE, estiloTitulo);

            // 4) Encabezados + anchos
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado horarios x detalles)
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<HorarioDTO> horarios = request.getHorarios();
            if (horarios != null) {
                for (HorarioDTO h : horarios) {
                    if (h == null) continue;
                    List<DetalleHorarioDTO> dets = h.getDetalles();
                    if (dets == null || dets.isEmpty()) continue;

                    for (DetalleHorarioDTO d : dets) {
                        if (d == null) continue;
                        Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                        UtilExcel.establecerValor(r, 0, item++, null);                          // ITEM
                        UtilExcel.establecerTexto(r, 1, nvl(h.getNombre()), null);              // HORARIO
                        UtilExcel.establecerTexto(r, 2, nvl(h.getCodigo()), null);              // CÓDIGO
                        UtilExcel.establecerTexto(r, 3, nvl(h.getHoraTrabajo()), null);         // HORAS DE TRABAJO
                        UtilExcel.establecerTexto(r, 4, nvl(h.getMinutosComida()), null);       // MINUTOS DE ALIMENTACIÓN
                        UtilExcel.establecerTexto(r, 5, h.isNoturno() ? "Sí" : "No", null);     // HORARIO NOTURNO
                        UtilExcel.establecerTexto(r, 6, nvl(h.getDocumento()), null);           // DOCUMENTO
                        UtilExcel.establecerValor(r, 7, d.getOrden(), null);                    // ORDEN
                        UtilExcel.establecerTexto(r, 8, nvl(d.getHora()), null);                // HORA
                        UtilExcel.establecerTexto(r, 9, nvl(d.getTolerancia()), null);          // TOLERANCIA
                        UtilExcel.establecerTexto(r,10, nvl(d.getTipoAccionShow()), null);      // ACCIÓN
                        UtilExcel.establecerTexto(r,11, d.isSegundoDia() ? "Sí" : "No", null);  // OTRO DÍA
                        UtilExcel.establecerValor(r,12, d.getMinutosAntes(), null);             // MINUTOS ANTES
                        UtilExcel.establecerValor(r,13, d.getMinutosDespues(), null);           // MINUTOS DESPUÉS
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                // Resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde, true);

                // 7) Tabla estilizada + filtros (ITEM sin filtro)
                boolean[] filtros = new boolean[HEADERS.length];
                for (int i = 0; i < filtros.length; i++) filtros[i] = true;
                filtros[0] = false;

                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "HorariosTabla",
                    FILA_ENCABEZADO, 0,
                    ultimaFila, HEADERS.length - 1,
                    true,
                    filtros
                );
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Horarios.xlsx", e); // Interno → 500
        }
    }

    
    // =========================
    //            CSV
    // =========================
    public byte[] generarReporteCSV(ReporteHorariosRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Horarios.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = {
            "n", "horario", "codigo", "horas_trabajo", "minutos_alimentacion", "horario_noturno",
            "documento", "orden", "hora", "tolerancia", "accion", "otro_dia",
            "minutos_antes", "minutos_despues"
        };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            int n = 1;
            List<HorarioDTO> horarios = request.getHorarios();
            if (horarios != null && !horarios.isEmpty()) {
                for (HorarioDTO h : horarios) {
                    List<DetalleHorarioDTO> dets = h.getDetalles();
                    if (dets == null || dets.isEmpty()) continue;

                    for (DetalleHorarioDTO d : dets) {
                        String horario        = (h.getNombre() == null)            ? "" : h.getNombre();
                        String codigo         = (h.getCodigo() == null)            ? "" : h.getCodigo();
                        String horasTrabajo   = (h.getHoraTrabajo() == null)       ? "" : String.valueOf(h.getHoraTrabajo());
                        String minutosAliment = (h.getMinutosComida() == null)     ? "" : String.valueOf(h.getMinutosComida());
                        String noturno        = h.isNoturno() ? "Sí" : "No";
                        String documento      = (h.getDocumento() == null)         ? "" : h.getDocumento();

                        String orden          = (d.getOrden() == null)             ? "" : String.valueOf(d.getOrden());
                        String hora           = (d.getHora() == null)              ? "" : String.valueOf(d.getHora());
                        String tolerancia     = (d.getTolerancia() == null)        ? "" : String.valueOf(d.getTolerancia());
                        String accion         = (d.getTipoAccionShow() == null)    ? "" : d.getTipoAccionShow();
                        String otroDia        = d.isSegundoDia() ? "Sí" : "No";
                        String minutosAntes   = (d.getMinutosAntes() == null)      ? "" : String.valueOf(d.getMinutosAntes());
                        String minutosDespues = (d.getMinutosDespues() == null)    ? "" : String.valueOf(d.getMinutosDespues());

                        sb.append(UtilCsv.csvEscape(String.valueOf(n++))).append(DELIM)
                        .append(UtilCsv.csvEscape(horario)).append(DELIM)
                        .append(UtilCsv.csvEscape(codigo)).append(DELIM)
                        .append(UtilCsv.csvEscape(horasTrabajo)).append(DELIM)
                        .append(UtilCsv.csvEscape(minutosAliment)).append(DELIM)
                        .append(UtilCsv.csvEscape(noturno)).append(DELIM)
                        .append(UtilCsv.csvEscape(documento)).append(DELIM)
                        .append(UtilCsv.csvEscape(orden)).append(DELIM)
                        .append(UtilCsv.csvEscape(hora)).append(DELIM)
                        .append(UtilCsv.csvEscape(tolerancia)).append(DELIM)
                        .append(UtilCsv.csvEscape(accion)).append(DELIM)
                        .append(UtilCsv.csvEscape(otroDia)).append(DELIM)
                        .append(UtilCsv.csvEscape(minutosAntes)).append(DELIM)
                        .append(UtilCsv.csvEscape(minutosDespues)).append(EOL);
                    }
                }
            }

            // Retorno (nunca null)
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Interno → 500
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    
    // =========================
    //            XML
    // =========================
    public byte[] generarReporteXML(ReporteHorariosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Horarios>\n");

            List<HorarioDTO> horarios = request.getHorarios();
            if (horarios != null) {
                for (HorarioDTO h : horarios) {
                    sb.append("  <horario codigo=\"").append(xml(h.getCodigo())).append("\">\n");
                    sb.append("    <nombre>").append(xml(h.getNombre())).append("</nombre>\n");
                    sb.append("    <horas_trabajo>").append(xml(h.getHoraTrabajo())).append("</horas_trabajo>\n");
                    sb.append("    <minutos_alimentación>").append(xml(h.getMinutosComida())).append("</minutos_alimentación>\n");
                    sb.append("    <horario_noturno>").append(h.isNoturno() ? "Sí" : "No").append("</horario_noturno>\n");
                    sb.append("    <documento>").append(xml(h.getDocumento())).append("</documento>\n");

                    List<DetalleHorarioDTO> dets = h.getDetalles();
                    if (dets != null && !dets.isEmpty()) {
                        sb.append("    <detalles>\n");
                        for (DetalleHorarioDTO d : dets) {
                            sb.append("      <detalle orden=\"").append(xml(d.getOrden())).append("\">\n");
                            sb.append("        <hora>").append(xml(d.getHora())).append("</hora>\n");
                            sb.append("        <tolerancia>").append(xml(d.getTolerancia())).append("</tolerancia>\n");
                            sb.append("        <accion>").append(xml(d.getTipoAccionShow())).append("</accion>\n");
                            sb.append("        <otro_dia>").append(d.isSegundoDia() ? "Sí" : "No").append("</otro_dia>\n");
                            sb.append("        <minutos_antes>").append(xml(d.getMinutosAntes())).append("</minutos_antes>\n");
                            sb.append("        <minutos_despues>").append(xml(d.getMinutosDespues())).append("</minutos_despues>\n");
                            sb.append("      </detalle>\n");
                        }
                        sb.append("    </detalles>\n");
                    }

                    sb.append("  </horario>\n");
                }
            }

            sb.append("</Horarios>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //        Helpers locales
    // =========================
    private String nvl(Object v) {
        return v == null ? "" : String.valueOf(v);
    }


    private String xml(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"","&quot;")
                .replace("'","&apos;");
    }
}

