package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Titulo.ReporteTitulosRequest;
import com.casapazmino.microservicio_reportes.model.Titulo.TituloDTO;
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
import java.util.List;

@Service
public class ReporteTituloService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteTitulosPDF(ReporteTitulosRequest request) {
        // DRY: constantes locales
        final float[] WIDTHS_TABLA       = { 2f, 4f, 6f };
        final float   PORCENTAJE_ANCHO   = 60f;
        final String  TITULO_REPORTE     = "LISTA DE TÍTULOS PROFESIONALES";

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

            // 2) Construcción (solo helpers existentes; no cambiamos diseño)
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Encabezados
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO_REPORTE));

            // Tabla principal
            PdfPTable tabla = new PdfPTable(3);
            tabla.setWidthPercentage(PORCENTAJE_ANCHO);
            tabla.setWidths(WIDTHS_TABLA);
            tabla.setSpacingBefore(10f);

            // Encabezados de la tabla
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NIVEL",  ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NOMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Filas (zebra)
            boolean zebra = false;
            for (TituloDTO t : request.getTitulos()) {
                Color bg = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(String.valueOf(t.getId()), ReporteUtil.fuenteTablaData(), bg));
                tabla.addCell(ReporteUtil.crearCelda(t.getNivel(),               ReporteUtil.fuenteTablaData(), bg));
                tabla.addCell(ReporteUtil.crearCelda(t.getNombre(),              ReporteUtil.fuenteTablaData(), bg));
                zebra = !zebra;
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → que el controller decida si es 400
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Titulos.pdf", e);
        } finally {
            // Ciclo de recursos garantizado
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
    public byte[] generarReporteTitulosXLSX(ReporteTitulosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Títulos"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;        // fila 6 (idx 5)

        // MERGES exactos (B1:D1 ... B5:D5) => (row 0..4, col 1..3)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 3;

        final String[] HEADERS = { "ITEM", "CÓDIGO", "NIVEL", "TÍTULO" };
        final int[]    ANCHOS  = {    10,     20,      30,      30     };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] { false, true, true, true };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges B1:D1 ... B5:D5
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos (B1 empresa, B2 "Lista de Títulos")
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                hoja, 0, 1,
                UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                estiloTitulo
            ); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TÍTULOS", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (respetar el orden recibido)
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            List<TituloDTO> titulos = request.getTitulos();
            if (titulos != null) {
                for (TituloDTO t : titulos) {
                    Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                    UtilExcel.establecerValor(r, 0, item++, null);                                   // ITEM
                    UtilExcel.establecerValor(r, 1, t.getId(), null);                                // CÓDIGO
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(t.getNivel()), null);    // NIVEL
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(t.getNombre()), null);   // TÍTULO
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(
                hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                0, HEADERS.length - 1, estiloCentroBorde, true
            );

            // Cuerpo: col 0 centrada; col 1..3 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 3, estiloIzqBorde, true);

                // 7) Tabla estilizada (A6:Dn), zebra y AutoFilter (ITEM sin filtro)
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "TitulosTabla",
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
            throw new ReportBuildException("No se pudo generar Titulos.xlsx", e);
        }
    }

    
    // =========================
    // CSV (orden simple de keys)
    // =========================
    public byte[] generarReporteTitulosCSV(ReporteTitulosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados iguales a las keys del front: id, nivel, nombre
            sb.append("id,nivel,nombre\n");

            List<TituloDTO> items = request.getTitulos();
            if (items != null) {
                for (TituloDTO t : items) {
                    String id = t.getId() == null ? "" : String.valueOf(t.getId());
                    String nivel = t.getNivel() == null ? "" : t.getNivel();
                    String nombre = t.getNombre() == null ? "" : t.getNombre();
                    sb.append(csv(id)).append(',')
                            .append(csv(nivel)).append(',')
                            .append(csv(nombre)).append('\n');
                }
            }

            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // XML (igual a xml2js del front)
    // =========================
    public byte[] generarReporteTitulosXML(ReporteTitulosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Titulos>\n");

            List<TituloDTO> items = request.getTitulos();
            if (items != null) {
                for (TituloDTO t : items) {
                    sb.append("  <titulos id=\"").append(xml(t.getId())).append("\">\n");
                    sb.append("    <nivel>").append(xml(t.getNivel())).append("</nivel>\n");
                    sb.append("    <nombre>").append(xml(t.getNombre())).append("</nombre>\n");
                    sb.append("  </titulos>\n");
                }
            }

            sb.append("</Titulos>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // Helpers CSV/XML
    // =========================
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
