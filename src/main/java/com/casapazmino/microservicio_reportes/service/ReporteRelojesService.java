package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Dispositivo.RelojDTO;
import com.casapazmino.microservicio_reportes.model.Dispositivo.ReporteRelojesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import org.openpdf.text.Document;
import org.openpdf.text.Font;
import org.openpdf.text.FontFactory;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;

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
public class ReporteRelojesService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteRelojesRequest request) {

        // ➊ DRY: constantes locales (no cambia look & feel)
        final String TITULO = "LISTA DE DISPOSITIVOS";
        final float[] WIDTHS = {
                2.2f, 3f, 2.2f, 4f, 4.5f, 2.5f, 3f, 2f, 2f, 2.5f, 2f, 4.5f, 3.5f, 3f, 4.5f
        };
        final String[] HEADERS = {
                "Código", "Empresa", "Ciudad", "Establecimiento", "Departamento", "Nombre", "IP",
                "Puerto", "Marca", "Modelo", "Serie", "Mac", "ID Fabricación", "Fabricante", "Zona Horaria"
        };
        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_BEFORE = 10f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate());
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
            document.add(ReporteUtil.crearTituloReporte(TITULO));

            // Tabla principal
            PdfPTable tabla = new PdfPTable(HEADERS.length);
            tabla.setWidthPercentage(WIDTH_PERCENT_100);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(SPACING_BEFORE);

            // Encabezados
            for (String h : HEADERS) {
                tabla.addCell(ReporteUtil.celdaEncabezadoTabla(h, COLOR_PRIMARIO));
            }

            // Cuerpo con zebra
            boolean zebra = false;
            List<RelojDTO> lista = request.getRelojes();
            if (lista != null) {
                for (RelojDTO r : lista) {
                    Color fondo = zebra ? COLOR_ZEBRA : Color.WHITE;
                    zebra = !zebra;

                    tabla.addCell(ReporteUtil.crearCelda(r.getCodigo(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomempresa(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomciudad(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomsucursal(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomdepar(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNombre(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getIp(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getPuerto()), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getMarca(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getModelo(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getSerie(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getMac(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getIdFabricacion(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getFabricante(), fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(
                            r.getZonaHorariaDispositivo() + " (" + r.getFormatoGmtDispositivo() + ")",
                            fuenteAuxiliarTabla(), fondo));
                }
            }

            document.add(tabla);

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → el controller decidirá 400 si aplica
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteRelojes.pdf", e);
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

    public Font fuenteAuxiliarTabla() {
        return FontFactory.getFont(FontFactory.HELVETICA, 7);
    }

    private String nv(Object v) {
        return v == null ? "" : String.valueOf(v);
    }

    // =========================
    // XLSX
    // =========================
    public byte[] generarReporteXLSX(ReporteRelojesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Relojes"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila visual 6 (idx 5)

        // MERGES exactos (B1:U1 ... B5:U5) => (row 0..4, col 1..20)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 20;

        final String[] HEADERS = {
                "ITEM", "ID", "CODIGO", "NOMBRE", "IP", "PUERTO", "CONTRASEÑA",
                "MARCA", "MODELO", "SERIE", "ID_FABRICACION", "FABRICANTE", "MAC",
                "TIPO CONEXION", "ID SUCURSAL", "ID DEPARTAMENTO",
                "NOMBRE DEPARTAMENTO", "NOMBRE CIUDAD", "TEMPERATURA",
                "ZONA HORARIA DISPOSITIVO", "FORMATO GMT DISPOSITIVO"
        };

        final int[] ANCHOS = {
                10, 20, 20, 20, 20, 15, 20, 20, 20, 20,
                20, 20, 20, 20, 20, 20, 25, 25, 15, 30, 30
        };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] {
                false, true, true, true, true, true, true,
                true, true, true, true, true, true, true,
                true, true, true, true, true, true, true
        };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:U1 ... B5:U5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                    hoja, 0, 1,
                    UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                    estiloTitulo); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE RELOJES", estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;

            List<RelojDTO> relojes = request.getRelojes();
            if (relojes != null) {
                for (int i = 0; i < relojes.size(); i++) {
                    RelojDTO r = relojes.get(i);
                    Row row = UtilExcel.asegurarFila(hoja, filaActual++);

                    UtilExcel.establecerValor(row, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(row, 1, r.getId(), null);
                    UtilExcel.establecerValor(row, 2, nv(r.getCodigo()), null);
                    UtilExcel.establecerValor(row, 3, nv(r.getNombre()), null);
                    UtilExcel.establecerValor(row, 4, nv(r.getIp()), null);
                    UtilExcel.establecerValor(row, 5, r.getPuerto(), null);
                    UtilExcel.establecerValor(row, 6, nv(r.getContrasenia()), null);
                    UtilExcel.establecerValor(row, 7, nv(r.getMarca()), null);
                    UtilExcel.establecerValor(row, 8, nv(r.getModelo()), null);
                    UtilExcel.establecerValor(row, 9, nv(r.getSerie()), null);
                    UtilExcel.establecerValor(row, 10, nv(r.getIdFabricacion()), null);
                    UtilExcel.establecerValor(row, 11, nv(r.getFabricante()), null);
                    UtilExcel.establecerValor(row, 12, nv(r.getMac()), null);
                    UtilExcel.establecerValor(row, 13, nv(r.getTipoConexion()), null);
                    UtilExcel.establecerValor(row, 14, r.getIdSucursal(), null);
                    UtilExcel.establecerValor(row, 15, r.getIdDepartamento(), null);
                    UtilExcel.establecerValor(row, 16, nv(r.getNomdepar()), null);
                    UtilExcel.establecerValor(row, 17, nv(r.getNomciudad()), null);
                    UtilExcel.establecerValor(row, 18, nv(r.getTemperatura()), null);
                    UtilExcel.establecerValor(row, 19, nv(r.getZonaHorariaDispositivo()), null);
                    UtilExcel.establecerValor(row, 20, nv(r.getFormatoGmtDispositivo()), null);
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) ALINEACIONES + BORDES (header centrado; cuerpo col 0 centrado, resto
            // izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(
                    hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            // Cuerpo: col 0 centrada; col 1..20 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde,
                        true);
            }

            // 7) TABLA estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "RelojesTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        FILTROS);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación → 400 (lo maneja el controller)
            throw e;
        } catch (Exception e) {
            // Interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Relojes.xlsx", e);
        }
    }

    // =========================
    // CSV (orden estable, similar a Excel)
    // =========================
    public byte[] generarReporteCSV(ReporteRelojesRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Relojes.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = {
                "id", "codigo", "nombre", "ip", "puerto", "contrasenia", "marca", "modelo", "serie",
                "id_fabricacion", "fabricante", "mac", "tipo_conexion", "id_sucursal", "id_departamento",
                "nomdepar", "nomciudad", "temperatura", "zona_horaria_dispositivo", "formato_gmt_dispositivo",
                "nomempresa", "nomsucursal"
        };

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
            List<RelojDTO> items = request.getRelojes();
            if (items != null && !items.isEmpty()) {
                for (RelojDTO r : items) {
                    String id = (r == null || r.getId() == null) ? "" : String.valueOf(r.getId());
                    String codigo = (r == null || r.getCodigo() == null) ? "" : r.getCodigo();
                    String nombre = (r == null || r.getNombre() == null) ? "" : r.getNombre();
                    String ip = (r == null || r.getIp() == null) ? "" : r.getIp();
                    String puerto = (r == null || r.getPuerto() == null) ? "" : String.valueOf(r.getPuerto());
                    String contrasenia = (r == null || r.getContrasenia() == null) ? "" : r.getContrasenia();
                    String marca = (r == null || r.getMarca() == null) ? "" : r.getMarca();
                    String modelo = (r == null || r.getModelo() == null) ? "" : r.getModelo();
                    String serie = (r == null || r.getSerie() == null) ? "" : r.getSerie();
                    String idFabricacion = (r == null || r.getIdFabricacion() == null) ? "" : r.getIdFabricacion();
                    String fabricante = (r == null || r.getFabricante() == null) ? "" : r.getFabricante();
                    String mac = (r == null || r.getMac() == null) ? "" : r.getMac();
                    String tipoConexion = (r == null || r.getTipoConexion() == null) ? "" : r.getTipoConexion();
                    String idSucursal = (r == null || r.getIdSucursal() == null) ? ""
                            : String.valueOf(r.getIdSucursal());
                    String idDepartamento = (r == null || r.getIdDepartamento() == null) ? ""
                            : String.valueOf(r.getIdDepartamento());
                    String nomdepar = (r == null || r.getNomdepar() == null) ? "" : r.getNomdepar();
                    String nomciudad = (r == null || r.getNomciudad() == null) ? "" : r.getNomciudad();
                    String temperatura = (r == null || r.getTemperatura() == null) ? "" : r.getTemperatura();
                    String zonaHorariaDisp = (r == null || r.getZonaHorariaDispositivo() == null) ? ""
                            : r.getZonaHorariaDispositivo();
                    String formatoGmtDisp = (r == null || r.getFormatoGmtDispositivo() == null) ? ""
                            : r.getFormatoGmtDispositivo();
                    String nomempresa = (r == null || r.getNomempresa() == null) ? "" : r.getNomempresa();
                    String nomsucursal = (r == null || r.getNomsucursal() == null) ? "" : r.getNomsucursal();

                    sb.append(UtilCsv.csvEscape(id)).append(DELIM)
                            .append(UtilCsv.csvEscape(codigo)).append(DELIM)
                            .append(UtilCsv.csvEscape(nombre)).append(DELIM)
                            .append(UtilCsv.csvEscape(ip)).append(DELIM)
                            .append(UtilCsv.csvEscape(puerto)).append(DELIM)
                            .append(UtilCsv.csvEscape(contrasenia)).append(DELIM)
                            .append(UtilCsv.csvEscape(marca)).append(DELIM)
                            .append(UtilCsv.csvEscape(modelo)).append(DELIM)
                            .append(UtilCsv.csvEscape(serie)).append(DELIM)
                            .append(UtilCsv.csvEscape(idFabricacion)).append(DELIM)
                            .append(UtilCsv.csvEscape(fabricante)).append(DELIM)
                            .append(UtilCsv.csvEscape(mac)).append(DELIM)
                            .append(UtilCsv.csvEscape(tipoConexion)).append(DELIM)
                            .append(UtilCsv.csvEscape(idSucursal)).append(DELIM)
                            .append(UtilCsv.csvEscape(idDepartamento)).append(DELIM)
                            .append(UtilCsv.csvEscape(nomdepar)).append(DELIM)
                            .append(UtilCsv.csvEscape(nomciudad)).append(DELIM)
                            .append(UtilCsv.csvEscape(temperatura)).append(DELIM)
                            .append(UtilCsv.csvEscape(zonaHorariaDisp)).append(DELIM)
                            .append(UtilCsv.csvEscape(formatoGmtDisp)).append(DELIM)
                            .append(UtilCsv.csvEscape(nomempresa)).append(DELIM)
                            .append(UtilCsv.csvEscape(nomsucursal)).append(EOL);
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
    // XML (igual a tu legacy front)
    // =========================
    public byte[] generarReporteXML(ReporteRelojesRequest request) {
        final String NOMBRE_REPORTE = "Relojes.xml";
        final String ROOT_TAG = "Relojes";
        final String ITEM_TAG = "reloj";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder();

            // 1) Encabezado
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            // 2) Cuerpo
            List<RelojDTO> items = (request == null) ? null : request.getRelojes();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (RelojDTO r : items) {
                    String zona = (((r.getZonaHorariaDispositivo() == null) ? "" : r.getZonaHorariaDispositivo())
                            + " ("
                            + ((r.getFormatoGmtDispositivo() == null) ? "" : r.getFormatoGmtDispositivo())
                            + ")").trim();

                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" id=\"").append(UtilXml.xmlEsc(r == null ? null : r.getId())).append("\">")
                            .append(EOL);

                    sb.append(IND).append(IND).append("<codigo>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getCodigo()))
                            .append("</codigo>").append(EOL);

                    sb.append(IND).append(IND).append("<nombre_empresa>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getNomempresa()))
                            .append("</nombre_empresa>").append(EOL);

                    sb.append(IND).append(IND).append("<nombre_ciudad>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getNomciudad()))
                            .append("</nombre_ciudad>").append(EOL);

                    sb.append(IND).append(IND).append("<nombre_sucursal>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getNomsucursal()))
                            .append("</nombre_sucursal>").append(EOL);

                    sb.append(IND).append(IND).append("<nombre_departamento>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getNomdepar()))
                            .append("</nombre_departamento>").append(EOL);

                    sb.append(IND).append(IND).append("<nombre>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getNombre()))
                            .append("</nombre>").append(EOL);

                    sb.append(IND).append(IND).append("<ip>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getIp()))
                            .append("</ip>").append(EOL);

                    sb.append(IND).append(IND).append("<puerto>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getPuerto()))
                            .append("</puerto>").append(EOL);

                    sb.append(IND).append(IND).append("<marca>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getMarca()))
                            .append("</marca>").append(EOL);

                    sb.append(IND).append(IND).append("<modelo>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getModelo()))
                            .append("</modelo>").append(EOL);

                    sb.append(IND).append(IND).append("<serie>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getSerie()))
                            .append("</serie>").append(EOL);

                    sb.append(IND).append(IND).append("<mac>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getMac()))
                            .append("</mac>").append(EOL);

                    sb.append(IND).append(IND).append("<id_fabricacion>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getIdFabricacion()))
                            .append("</id_fabricacion>").append(EOL);

                    sb.append(IND).append(IND).append("<fabricante>")
                            .append(UtilXml.xmlEsc(r == null ? null : r.getFabricante()))
                            .append("</fabricante>").append(EOL);

                    sb.append(IND).append(IND).append("<zona_horaria>")
                            .append(UtilXml.xmlEsc(zona))
                            .append("</zona_horaria>").append(EOL);

                    sb.append(IND).append("</").append(ITEM_TAG).append(">").append(EOL);
                }
            }

            // 3) Cierre
            sb.append("</").append(ROOT_TAG).append(">").append(EOL);
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

}
