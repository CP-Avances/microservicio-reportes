package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Nacionalidad.NacionalidadDTO;
import com.casapazmino.microservicio_reportes.model.Nacionalidad.ReporteNacionalidadesRequest;
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
import java.util.List;

@Service
public class ReporteNacionalidadesService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteNacionalidadesPDF(ReporteNacionalidadesRequest request) {

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
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (helpers existentes)
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE NACIONALIDADES"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // Tabla
            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(40);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(
                    ReporteUtil.crearCelda("NACIONALIDAD", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            List<NacionalidadDTO> lista = request.getNacionalidades();
            boolean zebra = false;
            if (lista != null) {
                for (NacionalidadDTO n : lista) {
                    Color fondo = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(
                            ReporteUtil.crearCelda(String.valueOf(n.getId()), ReporteUtil.fuenteTablaData(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(n.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // entrada inválida → controller la mapeará a 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteNacionalidades.pdf", e);
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

    // =========================
    // XLSX (igual al front)
    // =========================
    public byte[] generarReporteNacionalidadesXLSX(ReporteNacionalidadesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Nacionalidad"; // nombre exacto del front (≤31)
        final int FILA_ENCABEZADO = 5;

        // Merges B1:C1 ... B5:C5 => (row 0..4, col 1..2)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 2;

        final String TITULO_REPORTE = "LISTA DE NACIONALIDADES";
        final String[] HEADERS = { "ITEM", "CODIGO", "NACIONALIDAD" }; // sin tildes como el front
        final int[] ANCHOS = { 20, 30, 40 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener encabezado visible

            // 1) Logo estándar A1:B5 (si existe)
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

            // 5) Cuerpo (ordenar por id asc; nulls al final)
            List<NacionalidadDTO> data = request.getNacionalidades();
            List<NacionalidadDTO> ordenados = new java.util.ArrayList<>(data == null ? java.util.List.of() : data);
            ordenados.sort(java.util.Comparator.comparing(
                    NacionalidadDTO::getId,
                    java.util.Comparator.nullsLast(java.util.Comparator.naturalOrder())));

            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            for (NacionalidadDTO n : ordenados) {
                if (n == null)
                    continue;
                Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                UtilExcel.establecerValor(r, 0, item++, null); // ITEM
                UtilExcel.establecerValor(r, 1, n.getId(), null); // CODIGO
                UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(n.getNombre()), null);// NACIONALIDAD
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada; resto
            // izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 2, estiloIzqBorde, true);

                // 7) Tabla estilizada (nombre histórico del front) + filtros (ITEM sin filtro)
                boolean[] filtros = new boolean[] { false, true, true };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "NivelesTitulosTabla", // se mantiene el nombre usado en front
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        filtros);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Nacionalidades.xlsx", e); // Interno → 500
        }
    }

    // =========================
    // CSV (como el front: keys = id, nombre)
    // =========================
    public byte[] generarReporteNacionalidadesCSV(ReporteNacionalidadesRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Nacionalidades.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id", "nombre" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0)
                    sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<NacionalidadDTO> items = request.getNacionalidades();
            if (items != null && !items.isEmpty()) {
                for (NacionalidadDTO n : items) {
                    String id = (n == null || n.getId() == null) ? "" : String.valueOf(n.getId());
                    String nom = (n == null || n.getNombre() == null) ? "" : n.getNombre();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                            .append(UtilCsv.csvEscape(nom)).append(EOL);
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
    // XML (replica exacta del front, incl. raíz con tilde)
    // =========================
    public byte[] generarReporteNacionalidadesXML(ReporteNacionalidadesRequest request) {
        final String NOMBRE_REPORTE = "Nacionalidad.xml";
        final String ROOT_TAG = "Nacionalidades"; // se mantiene según contrato actual del front
        final String ITEM_TAG = "nacionalidad";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<NacionalidadDTO> items = request.getNacionalidades();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (NacionalidadDTO n : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(n.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<nacionalidad>")
                            .append(UtilXml.xmlEsc(n.getNombre()))
                            .append("</nacionalidad>").append(EOL);

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
