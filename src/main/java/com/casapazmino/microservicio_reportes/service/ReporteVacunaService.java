package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Vacuna.ReporteVacunasRequest;
import com.casapazmino.microservicio_reportes.model.Vacuna.VacunaDTO;
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
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@Service
public class ReporteVacunaService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteVacunasPDF(ReporteVacunasRequest request) {
        // DRY: constantes locales
        final float[] WIDTHS_TABLA     = { 2f, 6f };
        final float   PORC_ANCHO_TABLA = 45f;
        final String  TITULO_REPORTE   = "LISTA TIPOS DE VACUNAS";

        // Colores calculados una sola vez
        final Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color colorZebra     = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer  = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización de recursos PDF
            baos     = new ByteArrayOutputStream();
            document = new Document(PageSize.A4, 40, 40, 30, 50);
            writer   = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes; no cambiamos diseño)
            // Logo de empresa
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO_REPORTE));

            // Tabla principal
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(PORC_ANCHO_TABLA);
            tabla.setWidths(WIDTHS_TABLA);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NOMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Filas (zebra) – manteniendo el sort por id
            List<VacunaDTO> lista = request.getVacunas();
            lista.sort(Comparator.comparingInt(VacunaDTO::getId));
            boolean zebra = false;
            for (VacunaDTO v : lista) {
                Color bg = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(v.getId()), ReporteUtil.fuenteTablaData(), bg));
                tabla.addCell(ReporteUtil.crearCelda(v.getNombre(),               ReporteUtil.fuenteTablaData(), bg));
                zebra = !zebra;
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y lanza IAEx → que el controller decida 400
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Vacunas.pdf", e);
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
    public byte[] generarReporteVacunasXLSX(ReporteVacunasRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Vacunas"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;        // fila 6 (idx 5)

        // MERGES exactos (B1:C1 ... B5:C5) => (row 0..4, col 1..2)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 2;

        final String[] HEADERS = { "ITEM", "CODIGO", "NOMBRE" };
        final int[]    ANCHOS  = {    20,      30,      40   };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] { false, true, true };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:C1 ... B5:C5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                hoja, 0, 1,
                UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                estiloTitulo
            ); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA TIPOS DE VACUNAS", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Datos (ordenar por id de forma null-safe)
            List<VacunaDTO> datos = request.getVacunas();
            List<VacunaDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(v -> {
                Integer id = v.getId();
                return (id == null) ? Long.MIN_VALUE : id.longValue();
            }));

            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            for (VacunaDTO v : ordenados) {
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, item++, null);                               // ITEM
                UtilExcel.establecerValor(r, 1, v.getId(), null);                            // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(v.getNombre()), null); // NOMBRE
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado
            UtilExcel.aplicarEstiloARegion(
                hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                0, HEADERS.length - 1, estiloCentroBorde, true
            );

            // Cuerpo
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);    // CODIGO..NOMBRE

                // 7) Tabla estilizada (A6:Cn), zebra y AutoFilter (ITEM sin filtro)
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "VacunasTabla",
                    FILA_ENCABEZADO, 0,
                    ultimaFila, HEADERS.length - 1,
                    true,
                    FILTROS
                );
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones → 400
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar Vacunas.xlsx", e);
        }
    }

    
    // =========================
    // CSV
    // =========================
    public byte[] generarReporteVacunasCSV(ReporteVacunasRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Vacunas.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id", "nombre" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<VacunaDTO> items = request.getVacunas();
            if (items != null && !items.isEmpty()) {
                for (VacunaDTO v : items) {
                    String id     = (v == null || v.getId() == null)     ? "" : String.valueOf(v.getId());
                    String nombre = (v == null || v.getNombre() == null) ? "" : v.getNombre();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                    .append(UtilCsv.csvEscape(nombre)).append(EOL);
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
    public byte[] generarReporteVacunasXML(ReporteVacunasRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Vacunas>\n");

            List<VacunaDTO> datos = request.getVacunas();
            List<VacunaDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(this::safeLongId));

            for (VacunaDTO v : ordenados) {
                sb.append("  <vacuna id=\"").append(xml(String.valueOf(v.getId()))).append("\">\n");
                sb.append("    <nombre>").append(xml(v.getNombre())).append("</nombre>\n");
                sb.append("  </vacuna>\n");
            }

            sb.append("</Vacunas>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // Helpers
    // =========================
    private long safeLongId(VacunaDTO v) {
        try {
            return Long.parseLong(String.valueOf(v.getId()));
        } catch (Exception e) {
            return Long.MAX_VALUE;
        }
    }

    private String xml(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }
}
