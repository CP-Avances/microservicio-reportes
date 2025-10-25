package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Discapacidad.DiscapacidadDTO;
import com.casapazmino.microservicio_reportes.model.Discapacidad.ReporteDiscapacidadesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
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
public class ReporteDiscapacidadService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteDiscapacidadesPDF(ReporteDiscapacidadesRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 2f, 6f };

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
            document.add(ReporteUtil.crearTituloReporte("LISTA DE DISCAPACIDADES"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra     = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(40);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("NOMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            List<DiscapacidadDTO> lista = request.getDiscapacidades();
            boolean zebra = false;
            if (lista != null) {
                for (DiscapacidadDTO d : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(d.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(d.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y falla, que el controller lo maneje (posible 400)
            throw e;
        } catch (Exception e) {
            // Error interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteDiscapacidades.pdf", e);
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
    public byte[] generarReporteDiscapacidadesXLSX(ReporteDiscapacidadesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Discapacidades";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:C1 ... B5:C5 => (row 0..4, col 1..2)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 2;

        final String[] HEADERS = { "ITEM", "CODIGO", "NOMBRE" };
        final int[] ANCHOS      = {   20,     30,      40   };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges B1:C1 ... B5:C5
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE DISCAPACIDADES", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row header = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(header, c, HEADERS[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(header, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Ordenar por id y escribir cuerpo
            List<DiscapacidadDTO> datos = request.getDiscapacidades();
            List<DiscapacidadDTO> ordenados = new ArrayList<>(datos == null ? List.of() : datos);
            ordenados.sort(Comparator.comparingLong(d -> d.getId() == null ? Long.MAX_VALUE : d.getId()));

            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;

            for (int i = 0; i < ordenados.size(); i++) {
                DiscapacidadDTO d = ordenados.get(i);
                Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                UtilExcel.establecerValor(r, 0, i + 1, null);                                // ITEM
                UtilExcel.establecerValor(r, 1, d.getId(), null);                            // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(d.getNombre()), null); // NOMBRE
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 2, estiloIzqBorde, true);

                // 7) Tabla estilizada (A6:Cn), zebra y AutoFilter (ITEM sin filtro)
                boolean[] filtros = new boolean[] { false, true, true };
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "DiscapacidadesTabla",
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
            throw new ReportBuildException("No se pudo generar Discapacidades.xlsx", e); // Interno → 500
        }
    }

    
    // =========================
    // CSV (keys simples)
    // =========================
    public byte[] generarReporteDiscapacidadesCSV(ReporteDiscapacidadesRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Discapacidades.csv";
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
            List<DiscapacidadDTO> items = request.getDiscapacidades();
            if (items != null && !items.isEmpty()) {
                for (DiscapacidadDTO d : items) {
                    String id     = (d.getId() == null)     ? "" : String.valueOf(d.getId());
                    String nombre = (d.getNombre() == null) ? "" : d.getNombre();

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
    // XML (igual a xml2js: raíz y nodos)
    // =========================
    public byte[] generarReporteDiscapacidadesXML(ReporteDiscapacidadesRequest request) {
        final String NOMBRE_REPORTE = "Discapacidades.xml";
        final String ROOT_TAG = "Discapacidades";
        final String ITEM_TAG = "discapacidad";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<DiscapacidadDTO> datos = request.getDiscapacidades();
            if (datos == null || datos.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                List<DiscapacidadDTO> ordenados = new ArrayList<>(datos);
                ordenados.sort(Comparator.comparingLong(d ->
                    d.getId() == null ? Long.MAX_VALUE : d.getId()
                ));

                for (DiscapacidadDTO d : ordenados) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                    .append(" id=\"").append(UtilXml.xmlEsc(d.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<nombre>")
                    .append(UtilXml.xmlEsc(d.getNombre()))
                    .append("</nombre>").append(EOL);

                    sb.append(IND).append("</").append(ITEM_TAG).append(">").append(EOL);
                }
            }

            sb.append("</").append(ROOT_TAG).append(">").append(EOL);
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }


}
