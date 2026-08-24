package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ConfiguracionHoraExtra.ConfiguracionHoraExtraDTO;
import com.casapazmino.microservicio_reportes.model.ConfiguracionHoraExtra.ReporteConfiguracionHorasExtraRequest;
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
public class ReporteConfiguracionHorasExtraService {

    // ===================== PDF =====================
    public byte[] generarReporteConfiguracionHorasExtraPDF(ReporteConfiguracionHorasExtraRequest request) {
        final String TITULO_REPORTE = "CONFIGURACIÓN DE TIPOS DE HORAS EXTRA";
        final String[] HEADERS = {
                "DESCRIPCIÓN",
                "RÉGIMEN",
                "TIPO HORA EXTRA",
                "RECARGO",
                "ANTES JORNADA",
                "DESPUÉS JORNADA",
                "MIN. MÍNIMOS",
                "COMPENSACIÓN ATRASOS",
                "DESCONTAR PERMISOS",
                "APLICA EN",
                "HORARIO",
                "DESCONTAR COMIDA",
                "ESTADO"
        };

        final float[] WIDTHS = {
                2.1f, 2.0f, 1.9f, 1.3f, 1.3f, 1.3f,
                1.3f, 1.9f, 1.5f, 2.2f, 1.6f, 1.5f, 1.1f
        };

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

            for (ConfiguracionHoraExtraDTO he : request.getConfiguraciones()) {
                if (he == null) continue;

                tabla.addCell(ReporteUtil.celdaDataCentro(valor(he.getDescripcion()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(he.getNombreRegimen()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(he.getTipoHoraExtra()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(obtenerRecargo(he.getRecargoPorcentaje()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(he.getCalcularAntesJornada()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(he.getCalcularDespuesJornada()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(valor(he.getMinutosMinimosTrabajados()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(he.getPermiteCompensacionAtrasos()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(he.getDescontarPermisosDeExtras()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(obtenerAplicacion(he), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(obtenerHorario(he), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(siNo(he.getDescontarMinutosComida()), Color.WHITE));
                tabla.addCell(ReporteUtil.celdaDataCentro(obtenerEstado(he.getEstado()), Color.WHITE));
            }

            document.add(tabla);
            document.close();

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ConfiguracionHorasExtra.pdf", e);
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
    public byte[] generarReporteConfiguracionHorasExtraXLSX(ReporteConfiguracionHorasExtraRequest request) {
        final String NOMBRE_HOJA = "Configuracion HE";
        final String TITULO_REPORTE = "CONFIGURACIÓN DE TIPOS DE HORAS EXTRA";
        final int FILA_ENCABEZADO = 5;

        final String[] HEADERS = {
                "ITEM",
                "DESCRIPCIÓN",
                "RÉGIMEN",
                "TIPO HORA EXTRA",
                "RECARGO",
                "CALCULAR ANTES JORNADA",
                "CALCULAR DESPUÉS JORNADA",
                "MINUTOS MÍNIMOS TRABAJADOS",
                "PERMITE COMPENSACIÓN ATRASOS",
                "DESCONTAR PERMISOS",
                "APLICA EN",
                "HORARIO",
                "DESCONTAR MINUTOS COMIDA",
                "ESTADO"
        };

        final int[] ANCHOS = {
                10, 35, 30, 25, 15, 28, 30,
                30, 32, 24, 45, 25, 30, 15
        };

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

            for (ConfiguracionHoraExtraDTO he : request.getConfiguraciones()) {
                if (he == null) continue;

                Row fila = UtilExcel.asegurarFila(hoja, filaActual++);

                UtilExcel.establecerValor(fila, 0, item++, null);
                UtilExcel.establecerValor(fila, 1, valor(he.getDescripcion()), null);
                UtilExcel.establecerValor(fila, 2, valor(he.getNombreRegimen()), null);
                UtilExcel.establecerValor(fila, 3, valor(he.getTipoHoraExtra()), null);
                UtilExcel.establecerValor(fila, 4, obtenerRecargo(he.getRecargoPorcentaje()), null);
                UtilExcel.establecerValor(fila, 5, siNo(he.getCalcularAntesJornada()), null);
                UtilExcel.establecerValor(fila, 6, siNo(he.getCalcularDespuesJornada()), null);
                UtilExcel.establecerValor(fila, 7, he.getMinutosMinimosTrabajados(), null);
                UtilExcel.establecerValor(fila, 8, siNo(he.getPermiteCompensacionAtrasos()), null);
                UtilExcel.establecerValor(fila, 9, siNo(he.getDescontarPermisosDeExtras()), null);
                UtilExcel.establecerValor(fila, 10, obtenerAplicacion(he), null);
                UtilExcel.establecerValor(fila, 11, obtenerHorario(he), null);
                UtilExcel.establecerValor(fila, 12, siNo(he.getDescontarMinutosComida()), null);
                UtilExcel.establecerValor(fila, 13, obtenerEstado(he.getEstado()), null);
            }

            int ultimaFila = filaActual == filaDatosInicio
                    ? FILA_ENCABEZADO
                    : filaActual - 1;

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
                        1,
                        estiloCentroBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        2,
                        4,
                        estiloIzqBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        5,
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
                        true,
                        true,
                        true,
                        true
                };

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "ConfiguracionHorasExtraTabla",
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
            throw new ReportBuildException("No se pudo generar ConfiguracionHorasExtra.xlsx", e);
        }
    }

    // ===================== CSV =====================
    public byte[] generarReporteConfiguracionHorasExtraCSV(ReporteConfiguracionHorasExtraRequest request) {
        final String DELIM = ",";
        final String EOL = "\r\n";

        final String[] HEADERS = {
                "n",
                "descripcion",
                "regimen",
                "tipoHoraExtra",
                "recargo",
                "calcularAntesJornada",
                "calcularDespuesJornada",
                "minutosMinimosTrabajados",
                "permiteCompensacionAtrasos",
                "descontarPermisos",
                "aplicaEn",
                "horario",
                "descontarMinutosComida",
                "estado"
        };

        validarRequest(request);

        try {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) {
                    sb.append(DELIM);
                }
                sb.append(HEADERS[i]);
            }

            sb.append(EOL);

            int n = 1;

            for (ConfiguracionHoraExtraDTO he : request.getConfiguraciones()) {
                if (he == null) continue;

                sb.append(n++).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(he.getDescripcion()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(he.getNombreRegimen()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(he.getTipoHoraExtra()))).append(DELIM)
                        .append(UtilCsv.csvEscape(obtenerRecargo(he.getRecargoPorcentaje()))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(he.getCalcularAntesJornada()))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(he.getCalcularDespuesJornada()))).append(DELIM)
                        .append(UtilCsv.csvEscape(valor(he.getMinutosMinimosTrabajados()))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(he.getPermiteCompensacionAtrasos()))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(he.getDescontarPermisosDeExtras()))).append(DELIM)
                        .append(UtilCsv.csvEscape(obtenerAplicacion(he))).append(DELIM)
                        .append(UtilCsv.csvEscape(obtenerHorario(he))).append(DELIM)
                        .append(UtilCsv.csvEscape(siNo(he.getDescontarMinutosComida()))).append(DELIM)
                        .append(UtilCsv.csvEscape(obtenerEstado(he.getEstado())))
                        .append(EOL);
            }

            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ConfiguracionHorasExtra.csv", e);
        }
    }

    // ===================== XML =====================
    public byte[] generarReporteConfiguracionHorasExtraXML(ReporteConfiguracionHorasExtraRequest request) {
        final String EOL = "\n";
        final String IND = "  ";

        validarRequest(request);

        try {
            StringBuilder sb = new StringBuilder(8192);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<ConfiguracionesHorasExtra>").append(EOL);

            for (ConfiguracionHoraExtraDTO he : request.getConfiguraciones()) {
                if (he == null) continue;

                sb.append(IND)
                        .append("<configuracion id=\"")
                        .append(UtilXml.xmlEsc(he.getId()))
                        .append("\">")
                        .append(EOL);

                agregarNodoXml(sb, "descripcion", he.getDescripcion(), 2);
                agregarNodoXml(sb, "regimen", he.getNombreRegimen(), 2);
                agregarNodoXml(sb, "tipoHoraExtra", he.getTipoHoraExtra(), 2);
                agregarNodoXml(sb, "recargo", obtenerRecargo(he.getRecargoPorcentaje()), 2);
                agregarNodoXml(sb, "calcularAntesJornada", siNo(he.getCalcularAntesJornada()), 2);
                agregarNodoXml(sb, "calcularDespuesJornada", siNo(he.getCalcularDespuesJornada()), 2);
                agregarNodoXml(sb, "minutosMinimosTrabajados", he.getMinutosMinimosTrabajados(), 2);
                agregarNodoXml(sb, "permiteCompensacionAtrasos", siNo(he.getPermiteCompensacionAtrasos()), 2);
                agregarNodoXml(sb, "descontarPermisos", siNo(he.getDescontarPermisosDeExtras()), 2);
                agregarNodoXml(sb, "aplicaEn", obtenerAplicacion(he), 2);
                agregarNodoXml(sb, "horario", obtenerHorario(he), 2);
                agregarNodoXml(sb, "descontarMinutosComida", siNo(he.getDescontarMinutosComida()), 2);
                agregarNodoXml(sb, "estado", obtenerEstado(he.getEstado()), 2);

                sb.append(IND)
                        .append("</configuracion>")
                        .append(EOL);
            }

            sb.append("</ConfiguracionesHorasExtra>").append(EOL);

            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ConfiguracionHorasExtra.xml", e);
        }
    }

    // ===================== HELPERS =====================
    private void validarRequest(ReporteConfiguracionHorasExtraRequest request) {
        if (request == null) {
            throw new IllegalArgumentException("La solicitud del reporte es obligatoria.");
        }

        List<ConfiguracionHoraExtraDTO> configuraciones = request.getConfiguraciones();

        if (configuraciones == null || configuraciones.isEmpty()) {
            throw new IllegalArgumentException(
                    "No existen configuraciones de horas extra para generar el reporte."
            );
        }
    }

    private String valor(Object valor) {
        return valor == null ? "" : String.valueOf(valor);
    }

    private String siNo(Boolean valor) {
        return Boolean.TRUE.equals(valor) ? "SÍ" : "NO";
    }

    private String obtenerEstado(Boolean estado) {
        return Boolean.TRUE.equals(estado) ? "ACTIVO" : "INACTIVO";
    }

    private String obtenerRecargo(String recargo) {
        if (recargo == null || recargo.trim().isEmpty()) {
            return "";
        }

        return recargo + "%";
    }

    private String obtenerHorario(ConfiguracionHoraExtraDTO he) {
        String inicio = valor(he.getHoraInicio());
        String fin = valor(he.getHoraFinal());

        if (inicio.isEmpty() && fin.isEmpty()) {
            return "";
        }

        if (inicio.isEmpty()) {
            return fin;
        }

        if (fin.isEmpty()) {
            return inicio;
        }

        return inicio + " - " + fin;
    }

    private String obtenerAplicacion(ConfiguracionHoraExtraDTO he) {
        StringBuilder aplicacion = new StringBuilder();

        if (Boolean.TRUE.equals(he.getAplicaDiasLaborables())) {
            aplicacion.append("LABORABLES");
        }

        if (Boolean.TRUE.equals(he.getAplicaFinesSemana())) {
            if (aplicacion.length() > 0) {
                aplicacion.append(" / ");
            }

            aplicacion.append("FINES DE SEMANA");
        }

        if (Boolean.TRUE.equals(he.getAplicaFeriados())) {
            if (aplicacion.length() > 0) {
                aplicacion.append(" / ");
            }

            aplicacion.append("FERIADOS");
        }

        return aplicacion.length() == 0
                ? "NO APLICA"
                : aplicacion.toString();
    }

    private void agregarNodoXml(
            StringBuilder sb,
            String nodo,
            Object valor,
            int nivel
    ) {
        String indentacion = "  ".repeat(nivel);

        sb.append(indentacion)
                .append("<").append(nodo).append(">")
                .append(UtilXml.xmlEsc(valor))
                .append("</").append(nodo).append(">")
                .append("\n");
    }
}