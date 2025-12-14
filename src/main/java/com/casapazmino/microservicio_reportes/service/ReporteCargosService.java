package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Cargo.CargoDTO;
import com.casapazmino.microservicio_reportes.model.Cargo.ReporteCargosRequest;
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
public class ReporteCargosService {

    // =========================
    // PDF
    // =========================
    public byte[] generarReportePDF(ReporteCargosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 1f, 4f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4); // respetamos tu tamaño original
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
            document.add(ReporteUtil.crearTituloReporte("LISTA TIPO DE CARGOS"));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            PdfPTable tabla = new PdfPTable(2);
            tabla.setWidthPercentage(50); // respetamos tu diseño
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("ITEM", colorPrincipal));
            tabla.addCell(ReporteUtil.celdaEncabezadoTabla("CARGOS", colorPrincipal));

            // Cuerpo (zebra)
            boolean zebra = false;
            for (CargoDTO cargo : request.getCargos()) {
                Color fondo = zebra ? colorZebra : Color.WHITE;
                tabla.addCell(
                        ReporteUtil.celdaDataCentro(String.valueOf(cargo.getId()), fondo));
                tabla.addCell(ReporteUtil.celdaDataCentro(cargo.getCargo(), fondo));
                zebra = !zebra;
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // si algún helper valida y falla, dejamos que el controller lo trate (posible
            // 400)
            throw e;
        } catch (Exception e) {
            // fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar ReporteCargos.pdf", e);
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
    // XLSX (Como ExcelJS del front)
    // =========================
    public byte[] generarReporteXLSX(ReporteCargosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Tipos de Cargos";
        final int FILA_ENCABEZADO = 5;

        // Merges exactos (B1:C1 ... B5:C5) => (row 0..4, col 1..2)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 2;

        final String[] HEADERS = { "ITEM", "CÓDIGO", "CARGO" };
        final int[] ANCHOS = { 10, 30, 45 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) Merges exactos B1:C1 ... B5:C5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos en B1 y B2
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TIPOS DE CARGOS", estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (datos = [n++, id, cargo])
            int filaDatosIni = FILA_ENCABEZADO + 1; // 6 → idx 6
            int filaAct = filaDatosIni;
            int item = 1;

            List<CargoDTO> cargos = request.getCargos();
            if (cargos != null) {
                for (CargoDTO c : cargos) {
                    Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                    UtilExcel.establecerValor(r, 0, item++, null); // ITEM (secuencial)
                    UtilExcel.establecerValor(r, 1, c.getId(), null); // CÓDIGO (id)
                    UtilExcel.establecerValor(r, 2, UtilExcel.nuloComoVacio(c.getCargo()), null); // CARGO (se mantiene
                                                                                                  // helper actual)
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            // Cuerpo: col 0 centrado; col 1..2 izquierda
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 2, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada + filtros (ITEM sin filtro)
            if (ultimaFila >= filaDatosIni) {
                boolean[] filtros = new boolean[] { false, true, true };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TipoCargoTabla",
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
            throw new ReportBuildException("No se pudo generar TiposDeCargos.xlsx", e); // Interno → 500
        }
    }

    // =========================
    // CSV (idéntico al front)
    // =========================
    public byte[] generarReporteCSV(ReporteCargosRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Cargos.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "ITEM", "CARGOS" };

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
            List<CargoDTO> items = request.getCargos();
            if (items != null && !items.isEmpty()) {
                for (CargoDTO c : items) {
                    String item = (c.getId() == null) ? "" : String.valueOf(c.getId()); // contrato histórico: ITEM = id
                    String cargo = (c.getCargo() == null) ? "" : c.getCargo();

                    sb.append(UtilCsv.csvEscape(item)).append(DELIM).append(UtilCsv.csvEscape(cargo)).append(EOL);
                }
            }
            // Retorno (nunca null)
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (IllegalArgumentException e) {
            // Validación (el controller puede mapear a 400)
            throw e;
        } catch (Exception e) {
            // Error interno uniforme (→ 500 en controller)
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    // =========================
    // XML (idéntico a xml2js del front)
    // =========================
    public byte[] generarReporteXML(ReporteCargosRequest request) {
        final String NOMBRE_REPORTE = "Cargos.xml";
        final String ROOT_TAG = "Cargos";
        final String ITEM_TAG = "roles";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(4_096);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<CargoDTO> items = request.getCargos();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (CargoDTO c : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(c.getId())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<descripcion>")
                            .append(UtilXml.xmlEsc(c.getCargo()))
                            .append("</descripcion>").append(EOL);

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
