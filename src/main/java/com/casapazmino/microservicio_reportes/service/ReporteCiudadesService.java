package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Ciudad.CiudadDTO;
import com.casapazmino.microservicio_reportes.model.Ciudad.ReporteCiudadesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

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
public class ReporteCiudadesService {
    // =========================
    // PDF
    // =========================
    public byte[] generarReportePDF(ReporteCiudadesRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 2f, 4f };

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
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("LISTA DE CIUDADES"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.crearCelda("Provincia", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            tabla.addCell(ReporteUtil.crearCelda("Ciudad", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

            // Cuerpo (zebra)
            boolean zebra = false;
            for (CiudadDTO ciudad : request.getCiudades()) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(ReporteUtil.crearCelda(ciudad.getProvincia(), ReporteUtil.fuenteTablaData(), fondo));
                tabla.addCell(ReporteUtil.crearCelda(ciudad.getNombre(), ReporteUtil.fuenteTablaData(), fondo));
                zebra = !zebra;
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // si algún helper lanza IAEx, dejamos que el controller lo trate (posible 400)
            throw e;
        } catch (Exception e) {
            // 500 uniforme
            throw new ReportBuildException("No se pudo generar ReporteCiudades.pdf", e);
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
    // XLSX (idéntico al front)
    // =========================
    public byte[] generarReporteXLSX(ReporteCiudadesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Ciudades";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:E1 ... B5:E5 => (row 0..4, col 1..4)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 4;

        final String[] HEADERS = { "ITEM", "ID", "NOMBRE", "PROVINCIA", "ID_PROVINCIA" };
        final int[] ANCHOS = { 10, 20, 20, 20, 20 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges B1:E1 ... B5:E5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos en B1 y B2 (upper)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE CIUDADES", estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [index+1, id, nombre, provincia, id_prov])
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;

            List<CiudadDTO> ciudades = request.getCiudades();
            if (ciudades != null) {
                for (int i = 0; i < ciudades.size(); i++) {
                    CiudadDTO c = ciudades.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                    UtilExcel.establecerValor(r, 0, i + 1, null);
                    UtilExcel.establecerValor(r, 1, c.getId(), null);
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(c.getNombre()), null);
                    UtilExcel.establecerValor(r, 3, UtilExcel.nuloComoVacio(c.getProvincia()), null);
                    UtilExcel.establecerValor(r, 4, c.getId_prov(), null);
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..4 izquierda
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 4, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (A6:En), zebra y AutoFilter
            if (ultimaFila >= filaDatosIni) {
                boolean[] filtros = new boolean[] { false, true, true, true, true };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "CiudadesTabla",
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
            throw new ReportBuildException("No se pudo generar Ciudades.xlsx", e); // Interno → 500
        }
    }

    // =========================
    // CSV (dinámico del front → orden fijo aquí)
    // =========================
    public byte[] generarReporteCSV(ReporteCiudadesRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Ciudades.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "id", "nombre", "provincia", "id_prov" };

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
            List<CiudadDTO> items = request.getCiudades();
            if (items != null && !items.isEmpty()) {
                for (CiudadDTO c : items) {
                    String id = (c.getId() == null) ? "" : String.valueOf(c.getId());
                    String nombre = (c.getNombre() == null) ? "" : c.getNombre();
                    String prov = (c.getProvincia() == null) ? "" : c.getProvincia();
                    String idProv = (c.getId_prov() == null) ? "" : String.valueOf(c.getId_prov());

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                            .append(UtilCsv.csvEscape(nombre)).append(DELIM)
                            .append(UtilCsv.csvEscape(prov)).append(DELIM)
                            .append(UtilCsv.csvEscape(idProv)).append(EOL);
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
    // XML (igual a tu xml2js)
    // =========================
    public byte[] generarReporteXML(ReporteCiudadesRequest request) {
        final String NOMBRE_REPORTE = "Ciudades.xml";
        final String ROOT_TAG = "Ciudades";
        final String ITEM_TAG = "ciudad";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<CiudadDTO> items = request.getCiudades();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (CiudadDTO c : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(c.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<nombre>")
                            .append(UtilXml.xmlEsc(c.getNombre()))
                            .append("</nombre>").append(EOL);

                    sb.append(IND).append(IND).append("<provincia>")
                            .append(UtilXml.xmlEsc(c.getProvincia()))
                            .append("</provincia>").append(EOL);

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
