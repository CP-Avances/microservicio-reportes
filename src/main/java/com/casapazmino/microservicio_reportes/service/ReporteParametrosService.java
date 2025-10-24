package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Parametro.DetalleParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ParametroDTO;
import com.casapazmino.microservicio_reportes.model.Parametro.ReporteParametrosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.Document;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.List;

@Service
public class ReporteParametrosService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteParametrosPDF(ReporteParametrosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS_ENCABEZADO = { 6.9f, 2.1f };
        final float[] WIDTHS_DETALLES   = { 2f, 3f, 6f };
        final String[] HEADERS_DETALLES = { "CÓDIGO DETALLE", "DETALLE", "DESCRIPCIÓN" };

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
            document.add(ReporteUtil.crearTituloReporte("PARÁMETROS GENERALES"));

            // Colores
            Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

            // Iteración de parámetros
            if (request.getParametros() != null) {
                for (ParametroDTO parametro : request.getParametros()) {

                    // Encabezado del bloque de parámetro
                    PdfPTable encabezado = new PdfPTable(2);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(WIDTHS_ENCABEZADO);

                    PdfPCell celdaParametro = new PdfPCell(
                        new Phrase("PARÁMETRO: " + parametro.getDescripcion(), ReporteUtil.fuenteEncabezado())
                    );
                    celdaParametro.setBackgroundColor(colorPrincipal);
                    celdaParametro.setPaddingTop(2f);
                    celdaParametro.setPaddingBottom(3f);
                    celdaParametro.setPaddingLeft(5f);
                    celdaParametro.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                    encabezado.addCell(celdaParametro);

                    PdfPCell celdaCodigo = new PdfPCell(
                        new Phrase("CÓDIGO PARAMETRO: " + parametro.getId(), ReporteUtil.fuenteEncabezado())
                    );
                    celdaCodigo.setBackgroundColor(colorPrincipal);
                    celdaCodigo.setPaddingTop(2f);
                    celdaCodigo.setPaddingBottom(3f);
                    celdaCodigo.setPaddingRight(3f);
                    celdaCodigo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                    encabezado.addCell(celdaCodigo);

                    encabezado.setSpacingAfter(3f);
                    encabezado.setSpacingBefore(12.5f);
                    document.add(encabezado);

                    // Detalles (si existen)
                    if (parametro.getDetalles() != null && !parametro.getDetalles().isEmpty()) {
                        PdfPTable tabla = new PdfPTable(3);
                        tabla.setWidthPercentage(100);
                        tabla.setWidths(WIDTHS_DETALLES);

                        // Encabezados de detalles
                        for (String h : HEADERS_DETALLES) {
                            tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
                        }

                        // Cuerpo de detalles (diseño original en blanco)
                        for (DetalleParametroDTO d : parametro.getDetalles()) {
                            tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()),      ReporteUtil.fuenteTablaData(), Color.WHITE));
                            tabla.addCell(ReporteUtil.crearCelda(d.getDescripcion(),              ReporteUtil.fuenteTablaData(), Color.WHITE));
                            tabla.addCell(ReporteUtil.crearCelda(d.getObservacion(),              ReporteUtil.fuenteTablaData(), Color.WHITE));
                        }

                        document.add(tabla);
                    }
                }
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla → controller puede mapear 400
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteParametros.pdf", e);
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

    // METODO QUE GENERA EL XLSX
    public byte[] generarReporteParametrosXLSX(ReporteParametrosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Parametros"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;

        // Merges B1:D1 ... B5:D5 => (row 0..4, col 1..3)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 3;

        final String TITULO_REPORTE = "LISTA DE PARÁMETROS GENERALES";
        final String[] HEADERS = { "ITEM", "CÓDIGO PARAMETRO", "PARÁMETRO", "DETALLE", "DESCRIPCIÓN" };
        final int[]    ANCHOS  = { 10, 25, 50, 20, 160 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener encabezado visible

            // 1) LOGO estándar A1:B5 (si existe)
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) MERGES B1:D1 ... B5:D5
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS (B1 empresa, B2 subtítulo)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, TITULO_REPORTE, estiloTitulo);                                     // B2

            // 4) ENCABEZADOS + ANCHOS
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO (paridad con ExcelJS)
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<ParametroDTO> parametros = request.getParametros();
            if (parametros != null) {
                for (ParametroDTO p : parametros) {
                    if (p == null) continue;
                    boolean tieneDetalles = (p.getDetalles() != null && !p.getDetalles().isEmpty());
                    if (tieneDetalles) {
                        for (DetalleParametroDTO d : p.getDetalles()) {
                            if (d == null) continue;
                            Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                            UtilExcel.establecerValor(r, 0, item++, null);                                // ITEM
                            UtilExcel.establecerValor(r, 1, p.getId(), null);                             // CÓDIGO
                            UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(p.getDescripcion()), null);
                            UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(d.getDescripcion()), null);
                            UtilExcel.establecerValor(r, 4, UtilExcel.nuloComoVacio(d.getObservacion()), null);
                        }
                    } else {
                        Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                        UtilExcel.establecerValor(r, 0, item++, null);
                        UtilExcel.establecerValor(r, 1, p.getId(), null);
                        UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(p.getDescripcion()), null);
                        UtilExcel.establecerValor(r, 3, "", null);
                        UtilExcel.establecerValor(r, 4, "", null);
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) ALINEACIONES + BORDES
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentroBorde, true);

            // Cuerpo: col 0 y 1 centrado; 2..4 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 1, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 2, 4, estiloIzqBorde, true);

                // 7) TABLA estilizada + filtros (ITEM y CÓDIGO sin filtro)
                boolean[] filtros = new boolean[] { false, false, true, true, true };
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "ParametrosTabla",
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
            throw new ReportBuildException("No se pudo generar Parametros.xlsx", e); // Interno → 500
        }
    }

    
    public byte[] generarReporteParametrosCSV(ReporteParametrosRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Parametros.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "n", "codigoParametro", "parametro", "detalle", "descripcion" };

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
            List<ParametroDTO> parametros = request.getParametros();
            if (parametros != null && !parametros.isEmpty()) {
                for (ParametroDTO p : parametros) {
                    String codParam = (p == null || p.getId() == null) ? "" : String.valueOf(p.getId());
                    String param    = (p == null || p.getDescripcion() == null) ? "" : p.getDescripcion();

                    List<DetalleParametroDTO> detalles = (p == null) ? null : p.getDetalles();
                    if (detalles != null && !detalles.isEmpty()) {
                        for (DetalleParametroDTO d : detalles) {
                            String det = (d == null || d.getDescripcion() == null) ? "" : d.getDescripcion();
                            String obs = (d == null || d.getObservacion() == null) ? "" : d.getObservacion();

                            sb.append(n++).append(DELIM)
                            .append(UtilCsv.csvEscape(codParam)).append(DELIM)
                            .append(UtilCsv.csvEscape(param)).append(DELIM)
                            .append(UtilCsv.csvEscape(det)).append(DELIM)
                            .append(UtilCsv.csvEscape(obs)).append(EOL);
                        }
                    } else {
                        // Fila sin detalle/descripcion
                        sb.append(n++).append(DELIM)
                        .append(UtilCsv.csvEscape(codParam)).append(DELIM)
                        .append(UtilCsv.csvEscape(param)).append(DELIM)
                        .append(UtilCsv.csvEscape("")).append(DELIM)
                        .append(UtilCsv.csvEscape("")).append(EOL);
                    }
                }
            }

            // Retorno (nunca null)
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Interno → 500
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    
    public byte[] generarReporteParametrosXML(ReporteParametrosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<ParametrosGenerales>\n");

            List<ParametroDTO> parametros = request.getParametros();
            if (parametros != null) {
                for (ParametroDTO p : parametros) {
                    sb.append("  <parametro codigo=\"").append(xmlEsc(p.getId())).append("\">\n");
                    sb.append("    <nombre>").append(xmlEsc(p.getDescripcion())).append("</nombre>\n");
                    sb.append("    <detalles>\n");

                    List<DetalleParametroDTO> detalles = p.getDetalles();
                    if (detalles != null) {
                        for (DetalleParametroDTO d : detalles) {
                            sb.append("      <detalle codigo=\"").append(xmlEsc(d.getId())).append("\">\n");
                            // OJO: En tu XML original el nodo hijo también se llama "detalle"
                            sb.append("        <detalle>").append(xmlEsc(d.getDescripcion())).append("</detalle>\n");
                            sb.append("        <descripcion>").append(xmlEsc(d.getObservacion()))
                                    .append("</descripcion>\n");
                            sb.append("      </detalle>\n");
                        }
                    }
                    sb.append("    </detalles>\n");
                    sb.append("  </parametro>\n");
                }
            }
            sb.append("</ParametrosGenerales>\n");

            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private String xmlEsc(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }


}
