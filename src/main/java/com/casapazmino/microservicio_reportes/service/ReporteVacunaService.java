package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Vacuna.ReporteVacunasRequest;
import com.casapazmino.microservicio_reportes.model.Vacuna.VacunaDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
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
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Vacunas");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // 2) Merges B1:C1 ... B5:C5 (3 columnas)
            for (int r = 0; r < 5; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, 1, 2);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA TIPOS DE VACUNAS", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = { "ITEM", "CODIGO", "NOMBRE" };
            int[] anchos = { 20, 30, 40 };

            Row header = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(header, c, encabezados[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(header, encabezados.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Datos (ordenar por id de forma segura)
            List<VacunaDTO> datos = request.getVacunas();
            List<VacunaDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(this::safeLongId));

            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;

            for (int i = 0; i < ordenados.size(); i++) {
                VacunaDTO v = ordenados.get(i);
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                UtilExcel.establecerValor(r, 1, v.getId(), null); // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(v.getNombre()), null); // NOMBRE
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);

                // 7) Tabla estilizada (A6:Cn), zebra y AutoFilter (ITEM sin filtro)
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "VacunasTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[] { false, true, true });
            }

            libro.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // CSV
    // =========================
    public byte[] generarReporteVacunasCSV(ReporteVacunasRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("id,nombre\n");

            List<VacunaDTO> items = request.getVacunas();
            if (items != null) {
                for (VacunaDTO v : items) {
                    String id = String.valueOf(v.getId());
                    String nombre = v.getNombre() == null ? "" : v.getNombre();
                    sb.append(csv(id)).append(',').append(csv(nombre)).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
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

    private String csv(String v) {
        if (v == null)
            return "";
        boolean quote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return quote ? "\"" + s + "\"" : s;
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
