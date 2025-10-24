package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.EstadoCivil.EstadoCivilDTO;
import com.casapazmino.microservicio_reportes.model.EstadoCivil.ReporteEstadosCivilRequest;
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
public class ReporteEstadoCivilService {

    // =========================
    // PDF
    // =========================
    public byte[] generarReportePDF(ReporteEstadosCivilRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 2f, 5f };

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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE ESTADOS CIVIL"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra     = ReporteUtil.colorZebraClaro();

            // Tabla principal
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(30);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO",        ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("ESTADO CIVIL",  ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            List<EstadoCivilDTO> lista = request.getEstadosCivil();
            boolean zebra = false;
            if (lista != null) {
                for (EstadoCivilDTO e : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(e.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(e.getEstadoCivil(),         ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // si un helper valida y falla, que el controller lo maneje (posible 400)
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteEstadosCivil.pdf", e);
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
    
    // =========================
    // XLSX
    // =========================
    public byte[] generarReporteXLSX(ReporteEstadosCivilRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Estado Civil";       // ≤ 31 chars
        final int FILA_ENCABEZADO = 5;

        // Merges exactos (B1:C1 ... B5:C5) => (row 0..4, col 1..2)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 2;

        final String[] HEADERS = { "ITEM", "CODIGO", "ESTADO CIVIL" };   // labels exactos
        final int[]    ANCHOS  = { 20, 30, 40 };                         // anchos exactos

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar (A1:B5) - solo si hay imagen
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges exactos B1:C1 ... B5:C5
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos (B1 empresa, B2 título)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE ESTADOS CIVIL", estiloTitulo);                           // B2

            // 4) Encabezados + anchos
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (ordenar por id ASC como en el front)
            List<EstadoCivilDTO> estados = request.getEstadosCivil();
            List<EstadoCivilDTO> ordenados = new java.util.ArrayList<>(
                    estados == null ? java.util.List.of() : estados);
            ordenados.sort(java.util.Comparator.comparing(e -> e.getId() == null ? Integer.MAX_VALUE : e.getId()));

            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            for (EstadoCivilDTO e : ordenados) {
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, item++, null);                                      // ITEM
                UtilExcel.establecerValor(r, 1, e.getId(), null);                                   // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(e.getEstadoCivil()), null); // ESTADO CIVIL
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Estilos reutilizables
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..2 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);

                // 7) Tabla estilizada + filtros (ITEM sin filtro)
                boolean[] filtros = new boolean[] { false, true, true };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "EstadoCivilTabla",
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
            // Validación → el controller puede mapear a 400
            throw e;
        } catch (Exception e) {
            // Interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar EstadosCivil.xlsx", e);
        }
    }

    
    // =========================
    // CSV
    // =========================
    public byte[] generarReporteCSV(ReporteEstadosCivilRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "EstadosCivil.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id", "estado_civil" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<EstadoCivilDTO> items = request.getEstadosCivil();
            if (items != null && !items.isEmpty()) {
                for (EstadoCivilDTO e : items) {
                    String id   = (e.getId() == null)           ? "" : String.valueOf(e.getId());
                    String desc = (e.getEstadoCivil() == null)  ? "" : e.getEstadoCivil();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                    .append(UtilCsv.csvEscape(desc)).append(EOL);
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
    // XML
    // =========================
    public byte[] generarReporteXML(ReporteEstadosCivilRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Estados_Civil>\n"); // <- raíz con guion bajo

            List<EstadoCivilDTO> items = request.getEstadosCivil();
            if (items != null) {
                for (EstadoCivilDTO e : items) {
                    sb.append("  <estado_civil id=\"").append(xml(e.getId())).append("\">\n");
                    sb.append("    <estado_civil>").append(xml(e.getEstadoCivil())).append("</estado_civil>\n"); // <-
                                                                                                                 // etiqueta
                                                                                                                 // repetida
                    sb.append("  </estado_civil>\n");
                }
            }

            sb.append("</Estados_Civil>\n");
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    // =========================
    // Helpers CSV/XML
    // =========================

    private String xml(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }
}
