package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.TipoPermiso.ReporteTiposPermisosRequest;
import com.casapazmino.microservicio_reportes.model.TipoPermiso.TipoPermisoDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openpdf.text.Document;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

@Service
public class ReporteTiposPermisosService {

    // ===================== PDF =====================
    public byte[] generarReporteTiposPermisosPDF(ReporteTiposPermisosRequest request) {
        final String TITULO_REPORTE = "CONFIGURACIÓN DE TIPOS DE PERMISOS";
        final String[] HEADERS = {
                "DESCRIPCIÓN",
                "SUCURSAL",
                "DÍAS MÁXIMOS",
                "HORAS MÁXIMAS",
                "DÍAS PREVIOS PARA SOLICITAR",
                "DÍAS PERMITIDOS HACIA ATRÁS",
                "JUSTIFICACIÓN",
                "DOCUMENTO",
                "DESCUENTO",
                "INCLUYE ALIMENTACIÓN",
                "EMPLEADO PUEDE SOLICITAR"
        };
        final float[] WIDTHS = { 3.2f, 1.8f, 1.2f, 1.2f, 1.7f, 1.7f, 1.7f, 1.3f, 1.3f, 1.5f, 1.3f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            validarRequest(request);

            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate());
            writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO_REPORTE));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());

            PdfPTable tabla = new PdfPTable(HEADERS.length);
            tabla.setWidthPercentage(100);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);
            tabla.setHeaderRows(1);

            for (String header : HEADERS) {
                tabla.addCell(ReporteUtil.celdaEncabezadoTabla(header, colorPrincipal));
            }

            for (TipoPermisoDTO tp : request.getTiposPermisos()) {
                if (tp == null) continue;

                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getDescripcion()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getSucursalNombre()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getDiasMaximoPermiso()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getHorasMaximoPermiso()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getDiasAnticiparPermiso()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getCrearDiasAnteriores()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(obtenerJustificacion(tp), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(obtenerDocumento(tp.getDocumento()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(tp.getTipoDescuento()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(tp.getIncluirMinutosComida()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(tp.getSolicitaEmpleado()), Color.WHITE));
            }

            document.add(tabla);
            document.close();

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar TiposPermisos.pdf", e);
        } finally {
            if (document != null && document.isOpen()) {
                try { document.close(); } catch (Exception ignore) { }
            }
            if (writer != null) {
                try { writer.close(); } catch (Exception ignore) { }
            }
            if (baos != null) {
                try { baos.close(); } catch (Exception ignore) { }
            }
        }
    }

    // ===================== XLSX =====================
    public byte[] generarReporteTiposPermisosXLSX(ReporteTiposPermisosRequest request) {
        final String NOMBRE_HOJA = "Tipos Permisos";
        final String TITULO_REPORTE = "CONFIGURACIÓN DE TIPOS DE PERMISOS";
        final int FILA_ENCABEZADO = 5;

        final String[] HEADERS = {
                "ITEM",
                "DESCRIPCIÓN",
                "SUCURSAL",
                "DÍAS MÁXIMOS",
                "HORAS MÁXIMAS",
                "DÍAS PREVIOS PARA SOLICITAR",
                "DÍAS PERMITIDOS HACIA ATRÁS",
                "JUSTIFICACIÓN",
                "DOCUMENTO",
                "DESCUENTO",
                "INCLUYE ALIMENTACIÓN",
                "EMPLEADO PUEDE SOLICITAR"
        };

        final int[] ANCHOS = { 10, 45, 30, 18, 20, 30, 32, 24, 22, 20, 26, 28 };

        validarRequest(request);

        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            for (int r = 0; r <= 4; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, 1, HEADERS.length - 1);
            }

            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);

            UtilExcel.establecerTexto(
                    hoja,
                    0,
                    1,
                    UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    1,
                    1,
                    TITULO_REPORTE,
                    estiloTitulo
            );

            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);

            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }

            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            filaHeader.setHeightInPoints(18f);

            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            for (TipoPermisoDTO tp : request.getTiposPermisos()) {
                if (tp == null) continue;

                Row fila = UtilExcel.asegurarFila(hoja, filaActual++);

                UtilExcel.establecerValor(fila, 0, item++, null);
                UtilExcel.establecerValor(fila, 1, valor(tp.getDescripcion()), null);
                UtilExcel.establecerValor(fila, 2, valor(tp.getSucursalNombre()), null);
                UtilExcel.establecerValor(fila, 3, tp.getDiasMaximoPermiso(), null);
                UtilExcel.establecerValor(fila, 4, valor(tp.getHorasMaximoPermiso()), null);
                UtilExcel.establecerValor(fila, 5, tp.getDiasAnticiparPermiso(), null);
                UtilExcel.establecerValor(fila, 6, tp.getCrearDiasAnteriores(), null);
                UtilExcel.establecerValor(fila, 7, obtenerJustificacion(tp), null);
                UtilExcel.establecerValor(fila, 8, obtenerDocumento(tp.getDocumento()), null);
                UtilExcel.establecerValor(fila, 9, valor(tp.getTipoDescuento()), null);
                UtilExcel.establecerValor(fila, 10, siNo(tp.getIncluirMinutosComida()), null);
                UtilExcel.establecerValor(fila, 11, siNo(tp.getSolicitaEmpleado()), null);
            }

            int ultimaFila = filaActual == filaDatosInicio ? FILA_ENCABEZADO : filaActual - 1;

            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            UtilExcel.aplicarEstiloARegion(
                    hoja,
                    FILA_ENCABEZADO,
                    FILA_ENCABEZADO,
                    0,
                    HEADERS.length - 1,
                    estiloCentroBorde,
                    true
            );

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        0,
                        0,
                        estiloCentroBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        1,
                        2,
                        estiloIzqBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        3,
                        HEADERS.length - 1,
                        estiloCentroBorde,
                        true
                );

                boolean[] filtros = {
                        false,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true,
                        true
                };

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TiposPermisosTabla",
                        FILA_ENCABEZADO,
                        0,
                        ultimaFila,
                        HEADERS.length - 1,
                        true,
                        filtros
                );
            }

            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar TiposPermisos.xlsx", e);
        }
    }

    // ===================== CSV =====================
    public byte[] generarReporteTiposPermisosCSV(ReporteTiposPermisosRequest request) {
        final String DELIM = ",";
        final String EOL = "\r\n";

        final String[] HEADERS = {
                "n",
                "descripcion",
                "sucursal",
                "diasMaximos",
                "horasMaximas",
                "diasPreviosSolicitar",
                "diasPermitidosAtras",
                "justificacion",
                "documento",
                "descuento",
                "incluyeAlimentacion",
                "empleadoPuedeSolicitar"
        };

        validarRequest(request);

        try {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }

            sb.append(EOL);

            int n = 1;

            for (TipoPermisoDTO tp : request.getTiposPermisos()) {
                if (tp == null) continue;

                sb.append(n++).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getDescripcion()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getSucursalNombre()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getDiasMaximoPermiso()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getHorasMaximoPermiso()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getDiasAnticiparPermiso()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getCrearDiasAnteriores()))).append(DELIM)
                        .append(UtilCsv.csvEscape(obtenerJustificacion(tp))).append(DELIM)
                        .append(UtilCsv.csvEscape(obtenerDocumento(tp.getDocumento()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(tp.getTipoDescuento()))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(tp.getIncluirMinutosComida()))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(tp.getSolicitaEmpleado())))
                        .append(EOL);
            }

            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar TiposPermisos.csv", e);
        }
    }

    // ===================== XML =====================
    public byte[] generarReporteTiposPermisosXML(ReporteTiposPermisosRequest request) {
        final String EOL = "\n";
        final String IND = "  ";

        validarRequest(request);

        try {
            StringBuilder sb = new StringBuilder(8192);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<TiposPermisos>").append(EOL);

            for (TipoPermisoDTO tp : request.getTiposPermisos()) {
                if (tp == null) continue;

                sb.append(IND)
                        .append("<tipoPermiso id=\"")
                        .append(UtilXml.xmlEsc(tp.getId()))
                        .append("\">")
                        .append(EOL);

                agregarNodoXml(sb, "descripcion", tp.getDescripcion(), 2);
                agregarNodoXml(sb, "sucursal", tp.getSucursalNombre(), 2);
                agregarNodoXml(sb, "diasMaximos", tp.getDiasMaximoPermiso(), 2);
                agregarNodoXml(sb, "horasMaximas", tp.getHorasMaximoPermiso(), 2);
                agregarNodoXml(sb, "diasPreviosSolicitar", tp.getDiasAnticiparPermiso(), 2);
                agregarNodoXml(sb, "diasPermitidosAtras", tp.getCrearDiasAnteriores(), 2);
                agregarNodoXml(sb, "justificacion", obtenerJustificacion(tp), 2);
                agregarNodoXml(sb, "documento", obtenerDocumento(tp.getDocumento()), 2);
                agregarNodoXml(sb, "descuento", tp.getTipoDescuento(), 2);
                agregarNodoXml(sb, "incluyeAlimentacion", siNo(tp.getIncluirMinutosComida()), 2);
                agregarNodoXml(sb, "empleadoPuedeSolicitar", siNo(tp.getSolicitaEmpleado()), 2);

                sb.append(IND).append("</tipoPermiso>").append(EOL);
            }

            sb.append("</TiposPermisos>").append(EOL);

            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar TiposPermisos.xml", e);
        }
    }

    // ===================== HELPERS =====================
    private void validarRequest(ReporteTiposPermisosRequest request) {
        if (request == null) {
            throw new IllegalArgumentException("La solicitud del reporte es obligatoria.");
        }

        List<TipoPermisoDTO> tiposPermisos = request.getTiposPermisos();

        if (tiposPermisos == null || tiposPermisos.isEmpty()) {
            throw new IllegalArgumentException("No existen tipos de permisos para generar el reporte.");
        }
    }

    private String valor(Object valor) {
        return valor == null ? "" : String.valueOf(valor);
    }

    private String siNo(Boolean valor) {
        return Boolean.TRUE.equals(valor) ? "SÍ" : "NO";
    }

    private String obtenerDocumento(Boolean documento) {
        return Boolean.TRUE.equals(documento) ? "REQUERIDO" : "OPCIONAL";
    }

    private String obtenerJustificacion(TipoPermisoDTO tp) {
        if (!Boolean.TRUE.equals(tp.getJustificar())) {
            return "NO REQUIERE";
        }

        if (tp.getDiasJustificar() == null) {
            return "REQUIERE";
        }

        return tp.getDiasJustificar() + (tp.getDiasJustificar() == 1 ? " DÍA" : " DÍAS");
    }

    private void agregarNodoXml(StringBuilder sb, String nodo, Object valor, int nivel) {
        String indentacion = "  ".repeat(nivel);

        sb.append(indentacion)
                .append("<").append(nodo).append(">")
                .append(UtilXml.xmlEsc(valor))
                .append("</").append(nodo).append(">")
                .append("\n");
    }
}