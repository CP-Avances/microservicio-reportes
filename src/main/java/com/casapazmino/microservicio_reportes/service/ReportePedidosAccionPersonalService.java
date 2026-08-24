package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.PedidoAccionPersonal.PedidoAccionPersonalDTO;
import com.casapazmino.microservicio_reportes.model.PedidoAccionPersonal.ReportePedidosAccionPersonalRequest;
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
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class ReportePedidosAccionPersonalService {

    // ===================== PDF =====================
    public byte[] generarReportePedidosAccionPersonalPDF(ReportePedidosAccionPersonalRequest request) {
        final String TITULO_REPORTE = "PEDIDOS DE ACCIONES DE PERSONAL";

        final String[] HEADERS = {
                "CÓDIGO",
                "NÚMERO",
                "TIPO ACCIÓN",
                "DETALLE ACCIÓN",
                "EMPLEADO",
                "FECHA ELABORACIÓN",
                "RIGE DESDE",
                "RIGE HASTA",
                "VACACIÓN"
        };

        final float[] WIDTHS = {
                0.8f,
                1.3f,
                1.6f,
                2.3f,
                2.8f,
                1.5f,
                1.4f,
                1.4f,
                1.0f
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

            String logoBase64 = normalizarLogoBase64(request.getLogoBase64());
            Image logo = ReporteUtil.obtenerLogo(logoBase64);

            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO_REPORTE));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(
                    request.getColorPrincipal()
            );

            PdfPTable tabla = new PdfPTable(HEADERS.length);
            tabla.setWidthPercentage(100);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);
            tabla.setHeaderRows(1);

            for (String header : HEADERS) {
                tabla.addCell(
                        ReporteUtil.celdaEncabezadoTabla(
                                header,
                                colorPrincipal
                        )
                );
            }

            for (PedidoAccionPersonalDTO pedido : request.getPedidos()) {
                if (pedido == null) continue;

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        valor(pedido.getId()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        valor(pedido.getNumeroAccionPersonal()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        valor(pedido.getAccionPersonal()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        valor(pedido.getDetalleAccion()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        valor(pedido.getEmpleado()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        formatearFecha(pedido.getFechaElaboracion()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        formatearFecha(pedido.getFechaRigeDesde()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        formatearFecha(pedido.getFechaRigeHasta()),
                        Color.WHITE
                ));

                tabla.addCell(ReporteUtil.celdaDataCentro(
                        obtenerVacacion(pedido.getIdVacacion()),
                        Color.WHITE
                ));
            }

            document.add(tabla);
            document.close();

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar PedidosAccionPersonal.pdf",
                    e
            );
        } finally {
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

    // ===================== XLSX =====================
    public byte[] generarReportePedidosAccionPersonalXLSX(ReportePedidosAccionPersonalRequest request) {
        final String NOMBRE_HOJA = "Pedidos Accion Personal";
        final String TITULO_REPORTE = "PEDIDOS DE ACCIONES DE PERSONAL";
        final int FILA_ENCABEZADO = 5;

        final String[] HEADERS = {
                "ITEM",
                "CÓDIGO",
                "NÚMERO",
                "TIPO ACCIÓN",
                "DETALLE ACCIÓN",
                "EMPLEADO",
                "FECHA ELABORACIÓN",
                "RIGE DESDE",
                "RIGE HASTA",
                "VACACIÓN"
        };

        final int[] ANCHOS = {
                10,
                12,
                20,
                25,
                40,
                40,
                22,
                20,
                20,
                18
        };

        validarRequest(request);

        try (
                XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()
        ) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            String logoBase64 = normalizarLogoBase64(
                    request.getLogoBase64()
            );

            byte[] logo = UtilExcel.decodificarImagenBase64(
                    logoBase64
            );

            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(
                        libro,
                        hoja,
                        logo
                );
            }

            for (int r = 0; r <= 4; r++) {
                UtilExcel.combinarCeldas(
                        hoja,
                        r,
                        r,
                        1,
                        HEADERS.length - 1
                );
            }

            CellStyle estiloTitulo =
                    ConfiguracionExcel.crearEstiloTitulo(libro);

            UtilExcel.establecerTexto(
                    hoja,
                    0,
                    1,
                    UtilExcel.aMayusculasSeguras(
                            request.getEmpresa()
                    ),
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    1,
                    1,
                    TITULO_REPORTE,
                    estiloTitulo
            );

            Row filaHeader =
                    UtilExcel.asegurarFila(
                            hoja,
                            FILA_ENCABEZADO
                    );

            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(
                        filaHeader,
                        c,
                        HEADERS[c],
                        null
                );
            }

            CellStyle estiloEncabezado =
                    ConfiguracionExcel
                            .crearEstiloEncabezadoTabla(
                                    libro
                            );

            UtilExcel.aplicarEstiloAFila(
                    filaHeader,
                    HEADERS.length,
                    estiloEncabezado
            );

            UtilExcel.establecerAnchosColumnas(
                    hoja,
                    ANCHOS
            );

            filaHeader.setHeightInPoints(18f);

            int filaDatosInicio =
                    FILA_ENCABEZADO + 1;

            int filaActual =
                    filaDatosInicio;

            int item = 1;

            for (PedidoAccionPersonalDTO pedido : request.getPedidos()) {
                if (pedido == null) continue;

                Row fila =
                        UtilExcel.asegurarFila(
                                hoja,
                                filaActual++
                        );

                UtilExcel.establecerValor(
                        fila,
                        0,
                        item++,
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        1,
                        pedido.getId(),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        2,
                        valor(pedido.getNumeroAccionPersonal()),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        3,
                        valor(pedido.getAccionPersonal()),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        4,
                        valor(pedido.getDetalleAccion()),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        5,
                        valor(pedido.getEmpleado()),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        6,
                        formatearFecha(
                                pedido.getFechaElaboracion()
                        ),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        7,
                        formatearFecha(
                                pedido.getFechaRigeDesde()
                        ),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        8,
                        formatearFecha(
                                pedido.getFechaRigeHasta()
                        ),
                        null
                );

                UtilExcel.establecerValor(
                        fila,
                        9,
                        obtenerVacacion(
                                pedido.getIdVacacion()
                        ),
                        null
                );
            }

            int ultimaFila =
                    filaActual == filaDatosInicio
                            ? FILA_ENCABEZADO
                            : filaActual - 1;

            CellStyle estiloCentroBorde =
                    ConfiguracionExcel
                            .crearEstiloCentroConBorde(
                                    libro
                            );

            CellStyle estiloIzqBorde =
                    ConfiguracionExcel
                            .crearEstiloIzquierdaConBorde(
                                    libro
                            );

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
                        3,
                        estiloCentroBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        4,
                        5,
                        estiloIzqBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosInicio,
                        ultimaFila,
                        6,
                        HEADERS.length - 1,
                        estiloCentroBorde,
                        true
                );

                boolean[] filtros = {
                        false,
                        false,
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
                        "PedidosAccionPersonalTabla",
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
            throw new ReportBuildException(
                    "No se pudo generar PedidosAccionPersonal.xlsx",
                    e
            );
        }
    }

    // ===================== CSV =====================
    public byte[] generarReportePedidosAccionPersonalCSV(ReportePedidosAccionPersonalRequest request) {
        final String DELIM = ",";
        final String EOL = "\r\n";

        final String[] HEADERS = {
                "n",
                "codigo",
                "numero",
                "tipoAccion",
                "detalleAccion",
                "empleado",
                "fechaElaboracion",
                "rigeDesde",
                "rigeHasta",
                "vacacion"
        };

        validarRequest(request);

        try {
            StringBuilder sb =
                    new StringBuilder();

            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) {
                    sb.append(DELIM);
                }

                sb.append(HEADERS[i]);
            }

            sb.append(EOL);

            int n = 1;

            for (PedidoAccionPersonalDTO pedido : request.getPedidos()) {
                if (pedido == null) continue;

                sb.append(n++)
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                valor(pedido.getId())
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                valor(pedido.getNumeroAccionPersonal())
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                valor(pedido.getAccionPersonal())
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                valor(pedido.getDetalleAccion())
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                valor(pedido.getEmpleado())
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                formatearFecha(
                                        pedido.getFechaElaboracion()
                                )
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                formatearFecha(
                                        pedido.getFechaRigeDesde()
                                )
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                formatearFecha(
                                        pedido.getFechaRigeHasta()
                                )
                        ))
                        .append(DELIM)

                        .append(UtilCsv.csvEscape(
                                obtenerVacacion(
                                        pedido.getIdVacacion()
                                )
                        ))

                        .append(EOL);
            }

            return sb.toString()
                    .getBytes(
                            StandardCharsets.UTF_8
                    );

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar PedidosAccionPersonal.csv",
                    e
            );
        }
    }

    // ===================== XML =====================
    public byte[] generarReportePedidosAccionPersonalXML(ReportePedidosAccionPersonalRequest request) {
        final String EOL = "\n";
        final String IND = "  ";

        validarRequest(request);

        try {
            StringBuilder sb =
                    new StringBuilder(8192);

            sb.append(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            ).append(EOL);

            sb.append(
                    "<PedidosAccionPersonal>"
            ).append(EOL);

            for (PedidoAccionPersonalDTO pedido : request.getPedidos()) {
                if (pedido == null) continue;

                sb.append(IND)
                        .append("<pedido id=\"")
                        .append(
                                UtilXml.xmlEsc(
                                        pedido.getId()
                                )
                        )
                        .append("\">")
                        .append(EOL);

                agregarNodoXml(
                        sb,
                        "numero",
                        pedido.getNumeroAccionPersonal(),
                        2
                );

                agregarNodoXml(
                        sb,
                        "tipoAccion",
                        pedido.getAccionPersonal(),
                        2
                );

                agregarNodoXml(
                        sb,
                        "detalleAccion",
                        pedido.getDetalleAccion(),
                        2
                );

                agregarNodoXml(
                        sb,
                        "empleado",
                        pedido.getEmpleado(),
                        2
                );

                agregarNodoXml(
                        sb,
                        "fechaElaboracion",
                        formatearFecha(
                                pedido.getFechaElaboracion()
                        ),
                        2
                );

                agregarNodoXml(
                        sb,
                        "rigeDesde",
                        formatearFecha(
                                pedido.getFechaRigeDesde()
                        ),
                        2
                );

                agregarNodoXml(
                        sb,
                        "rigeHasta",
                        formatearFecha(
                                pedido.getFechaRigeHasta()
                        ),
                        2
                );

                agregarNodoXml(
                        sb,
                        "vacacion",
                        obtenerVacacion(
                                pedido.getIdVacacion()
                        ),
                        2
                );

                sb.append(IND)
                        .append("</pedido>")
                        .append(EOL);
            }

            sb.append(
                    "</PedidosAccionPersonal>"
            ).append(EOL);

            return sb.toString()
                    .getBytes(
                            StandardCharsets.UTF_8
                    );

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar PedidosAccionPersonal.xml",
                    e
            );
        }
    }

    // ===================== HELPERS =====================
    private void validarRequest(
            ReportePedidosAccionPersonalRequest request
    ) {

        if (request == null) {
            throw new IllegalArgumentException(
                    "La solicitud del reporte es obligatoria."
            );
        }

        List<PedidoAccionPersonalDTO> pedidos =
                request.getPedidos();

        if (pedidos == null || pedidos.isEmpty()) {
            throw new IllegalArgumentException(
                    "No existen pedidos de acciones de personal para generar el reporte."
            );
        }
    }

    private String valor(Object valor) {
        return valor == null
                ? ""
                : String.valueOf(valor);
    }

    private String normalizarLogoBase64(String logoBase64) {
        if (
                logoBase64 == null ||
                logoBase64.trim().isEmpty() ||
                "null".equalsIgnoreCase(
                        logoBase64.trim()
                )
        ) {
            return null;
        }

        return logoBase64;
    }

    private String obtenerVacacion(Integer idVacacion) {
        if (idVacacion == null) {
            return "—";
        }

        return String.valueOf(idVacacion);
    }

    private String formatearFecha(String fecha) {
        if (
                fecha == null ||
                fecha.trim().isEmpty()
        ) {
            return "";
        }

        try {
            String fechaIso =
                    fecha.length() >= 10
                            ? fecha.substring(0, 10)
                            : fecha;

            LocalDate localDate =
                    LocalDate.parse(fechaIso);

            return localDate.format(
                    DateTimeFormatter.ofPattern(
                            "dd/MM/yyyy"
                    )
            );

        } catch (Exception e) {
            return fecha;
        }
    }

    private void agregarNodoXml(
            StringBuilder sb,
            String nodo,
            Object valor,
            int nivel
    ) {

        String indentacion =
                "  ".repeat(nivel);

        sb.append(indentacion)
                .append("<")
                .append(nodo)
                .append(">")
                .append(
                        UtilXml.xmlEsc(valor)
                )
                .append("</")
                .append(nodo)
                .append(">")
                .append("\n");
    }
}