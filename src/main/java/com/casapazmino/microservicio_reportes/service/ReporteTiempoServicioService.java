package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio.ContratoTiempoServicioDTO;
import com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio.EmpleadoTiempoServicioDTO;
import com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio.GrupoTiempoServicioDTO;
import com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio.ReporteTiempoServicioRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.Phrase;
import org.openpdf.text.Rectangle;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Locale;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTiempoServicioService {

    private static final String NOMBRE_REPORTE =
            "REPORTE DE TIEMPO DE SERVICIO";

    private static final String NOMBRE_HOJA =
            "Tiempo_servicio";

    private static final int COLUMNAS_REPORTE = 10;

    private static final int FILA_ENCABEZADO_EXCEL = 5;

    /*
     * Columnas PDF:
     *
     * 1. N°
     * 2. Identificación
     * 3. Código
     * 4. Apellidos y nombres
     * 5. Departamento
     * 6. Cargo
     * 7. Régimen laboral
     * 8. Fecha ingreso
     * 9. Fecha salida
     * 10. Tiempo de servicio
     */
    private static final float[] ANCHOS_PDF = {
            0.45f,
            1.35f,
            0.75f,
            2.10f,
            1.75f,
            1.90f,
            1.55f,
            1.10f,
            1.10f,
            1.65f
    };

    private static final String[] ENCABEZADOS = {
            "N°",
            "IDENTIFICACIÓN",
            "CÓDIGO",
            "APELLIDOS Y NOMBRES",
            "DEPARTAMENTO",
            "CARGO",
            "RÉGIMEN LABORAL",
            "F. INGRESO",
            "F. SALIDA",
            "TIEMPO DE SERVICIO"
    };

    private static final int[] ANCHOS_EXCEL = {
            10,
            20,
            15,
            32,
            28,
            30,
            25,
            20,
            20,
            30
    };

    /*
     * La columna ITEM no tendrá filtro.
     * Las demás columnas sí.
     */
    private static final boolean[] FILTROS_EXCEL = {
            false,
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

    private static final Locale LOCALE_ES =
            new Locale("es", "ES");

    private static final DateTimeFormatter FORMATO_DIA =
            DateTimeFormatter.ofPattern("EEE", LOCALE_ES);

    private static final DateTimeFormatter FORMATO_FECHA =
            DateTimeFormatter.ofPattern("dd/MM/yyyy", LOCALE_ES);

    // =====================================================================
    // GENERACIÓN DEL PDF
    // =====================================================================

    public byte[] generarReporteTiempoServicioPDF(
            ReporteTiempoServicioRequest request
    ) {
        validarRequest(request);

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();

            document = new Document(
                    PageSize.A4.rotate(),
                    25,
                    25,
                    30,
                    50
            );

            writer = PdfWriter.getInstance(
                    document,
                    baos
            );

            writer.setPageEvent(
                    new ConfiguracionPaginaPDF(
                            request.getUsuario(),
                            request.getFraseMarcaAgua(),
                            request.getColorPrincipal()
                    )
            );

            document.open();

            final Color colorPrimario =
                    ReporteUtil.convertirHexAColor(
                            request.getColorPrincipal()
                    );

            final Color colorSecundario =
                    ReporteUtil.convertirHexAColor(
                            request.getColorSecundario()
                    );

            final Color colorZebra =
                    ReporteUtil.colorZebraClaro();

            agregarLogoPDF(
                    document,
                    request.getLogoBase64()
            );

            document.add(
                    ReporteUtil.crearTituloEmpresa(
                            safe(request.getEmpresa())
                    )
            );

            document.add(
                    ReporteUtil.crearTituloReporte(
                            construirTitulo(request)
                    )
            );

            document.add(
                    ReporteUtil.crearTituloPeriodo(
                            "FECHA DE CORTE: "
                                    + formatearFecha(
                                            request.getFechaCorte()
                                    )
                    )
            );

            int totalRegistros =
                    contarContratos(request);

            document.add(
                    crearTituloLista(
                            totalRegistros,
                            colorSecundario
                    )
            );

            PdfPTable tabla =
                    new PdfPTable(COLUMNAS_REPORTE);

            tabla.setWidthPercentage(100);
            tabla.setWidths(ANCHOS_PDF);
            tabla.setHeaderRows(1);
            tabla.setSplitLate(false);

            agregarEncabezadosPDF(
                    tabla,
                    colorPrimario
            );

            AtomicInteger numeroRegistro =
                    new AtomicInteger(1);

            if (request.getGrupos() != null) {
                for (
                        GrupoTiempoServicioDTO grupo
                        : request.getGrupos()
                ) {
                    agregarGrupoPDF(
                            tabla,
                            grupo,
                            numeroRegistro,
                            colorZebra
                    );
                }
            }

            tabla.setSpacingAfter(10f);

            document.add(tabla);

            document.close();

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;

        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar TiempoServicio.pdf",
                    e
            );

        } finally {
            if (
                    document != null &&
                    document.isOpen()
            ) {
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

    private void agregarLogoPDF(
            Document document,
            String logoBase64
    ) throws Exception {
        if (
                logoBase64 == null ||
                logoBase64.isBlank()
        ) {
            return;
        }

        Image logo =
                ReporteUtil.obtenerLogo(
                        logoBase64
                );

        if (logo != null) {
            document.add(logo);
        }
    }

    private void agregarEncabezadosPDF(
            PdfPTable tabla,
            Color colorPrimario
    ) {
        for (String encabezado : ENCABEZADOS) {
            PdfPCell celda =
                    ReporteUtil.crearCelda(
                            encabezado,
                            ReporteUtil.fuenteEncabezadoTablaData(),
                            colorPrimario
                    );

            celda.setHorizontalAlignment(
                    Element.ALIGN_CENTER
            );

            celda.setVerticalAlignment(
                    Element.ALIGN_MIDDLE
            );

            celda.setPadding(4f);

            tabla.addCell(celda);
        }
    }

    private void agregarGrupoPDF(
            PdfPTable tabla,
            GrupoTiempoServicioDTO grupo,
            AtomicInteger numeroRegistro,
            Color colorZebra
    ) {
        if (
                grupo == null ||
                grupo.getEmpleados() == null
        ) {
            return;
        }

        for (
                EmpleadoTiempoServicioDTO empleado
                : grupo.getEmpleados()
        ) {
            agregarEmpleadoPDF(
                    tabla,
                    empleado,
                    numeroRegistro,
                    colorZebra
            );
        }
    }

    private void agregarEmpleadoPDF(
            PdfPTable tabla,
            EmpleadoTiempoServicioDTO empleado,
            AtomicInteger numeroRegistro,
            Color colorZebra
    ) {
        if (
                empleado == null ||
                empleado.getTiempoServicio() == null
        ) {
            return;
        }

        String nombreCompleto =
                construirNombreEmpleado(empleado);

        for (
                ContratoTiempoServicioDTO contrato
                : empleado.getTiempoServicio()
        ) {
            if (contrato == null) {
                continue;
            }

            int numero =
                    numeroRegistro.getAndIncrement();

            Color fondo =
                    numero % 2 == 0
                            ? colorZebra
                            : Color.WHITE;

            agregarCeldaPDF(
                    tabla,
                    String.valueOf(numero),
                    fondo,
                    Element.ALIGN_CENTER
            );

            agregarCeldaPDF(
                    tabla,
                    safe(empleado.getIdentificacion()),
                    fondo,
                    Element.ALIGN_CENTER
            );

            agregarCeldaPDF(
                    tabla,
                    safe(empleado.getCodigo()),
                    fondo,
                    Element.ALIGN_CENTER
            );

            agregarCeldaPDF(
                    tabla,
                    nombreCompleto,
                    fondo,
                    Element.ALIGN_LEFT
            );

            agregarCeldaPDF(
                    tabla,
                    firstNonEmpty(
                            contrato.getDepartamento(),
                            empleado.getDepartamento()
                    ),
                    fondo,
                    Element.ALIGN_LEFT
            );

            agregarCeldaPDF(
                    tabla,
                    firstNonEmpty(
                            contrato.getCargo(),
                            empleado.getCargo()
                    ),
                    fondo,
                    Element.ALIGN_LEFT
            );

            agregarCeldaPDF(
                    tabla,
                    firstNonEmpty(
                            contrato.getRegimen(),
                            empleado.getRegimen()
                    ),
                    fondo,
                    Element.ALIGN_LEFT
            );

            agregarCeldaPDF(
                    tabla,
                    formatearFecha(
                            contrato.getFechaIngreso()
                    ),
                    fondo,
                    Element.ALIGN_CENTER
            );

            agregarCeldaPDF(
                    tabla,
                    formatearFechaOpcional(
                            contrato.getFechaSalida()
                    ),
                    fondo,
                    Element.ALIGN_CENTER
            );

            agregarCeldaPDF(
                    tabla,
                    obtenerTiempoServicio(contrato),
                    fondo,
                    Element.ALIGN_CENTER
            );
        }
    }

    private void agregarCeldaPDF(
            PdfPTable tabla,
            String texto,
            Color fondo,
            int alineacion
    ) {
        PdfPCell celda =
                ReporteUtil.crearCelda(
                        safe(texto),
                        ReporteUtil.fuenteTablaData(),
                        fondo
                );

        celda.setHorizontalAlignment(alineacion);

        celda.setVerticalAlignment(
                Element.ALIGN_MIDDLE
        );

        celda.setPadding(3f);

        tabla.addCell(celda);
    }

    private PdfPTable crearTituloLista(
            int totalRegistros,
            Color colorSecundario
    ) throws Exception {
        PdfPTable tablaTitulo =
                new PdfPTable(2);

        tablaTitulo.setWidthPercentage(100);

        tablaTitulo.setWidths(
                new float[]{
                        8f,
                        2f
                }
        );

        PdfPCell celdaTitulo =
                new PdfPCell(
                        new Phrase(
                                "LISTA DE EMPLEADOS",
                                ReporteUtil.fuenteEncabezadoTablaData()
                        )
                );

        celdaTitulo.setBackgroundColor(
                colorSecundario
        );

        celdaTitulo.setPadding(5f);

        celdaTitulo.setVerticalAlignment(
                Element.ALIGN_MIDDLE
        );

        celdaTitulo.setBorder(
                Rectangle.TOP
                        | Rectangle.BOTTOM
                        | Rectangle.LEFT
        );

        tablaTitulo.addCell(celdaTitulo);

        PdfPCell celdaContador =
                new PdfPCell(
                        new Phrase(
                                "N° Registros: "
                                        + totalRegistros,
                                ReporteUtil.fuenteEncabezadoTablaData()
                        )
                );

        celdaContador.setBackgroundColor(
                colorSecundario
        );

        celdaContador.setHorizontalAlignment(
                Element.ALIGN_RIGHT
        );

        celdaContador.setVerticalAlignment(
                Element.ALIGN_MIDDLE
        );

        celdaContador.setPadding(5f);

        celdaContador.setBorder(
                Rectangle.TOP
                        | Rectangle.BOTTOM
                        | Rectangle.RIGHT
        );

        tablaTitulo.addCell(celdaContador);

        tablaTitulo.setSpacingAfter(10f);

        return tablaTitulo;
    }

    // =====================================================================
    // GENERACIÓN DEL EXCEL
    // =====================================================================

    public byte[] generarReporteTiempoServicioExcel(
            ReporteTiempoServicioRequest request
    ) {
        validarRequest(request);

        try (
                XSSFWorkbook libro =
                        new XSSFWorkbook();

                ByteArrayOutputStream baos =
                        new ByteArrayOutputStream()
        ) {
            XSSFSheet hoja =
                    libro.createSheet(
                            NOMBRE_HOJA
                    );

            hoja.createFreezePane(
                    0,
                    FILA_ENCABEZADO_EXCEL + 1
            );

            agregarLogoExcel(
                    libro,
                    hoja,
                    request.getLogoBase64()
            );

            /*
             * El reporte tiene 10 columnas: A:J.
             *
             * Se combinan las celdas B:J para los títulos,
             * dejando la primera columna disponible para el logo.
             */
            for (int fila = 0; fila <= 4; fila++) {
                UtilExcel.combinarCeldas(
                        hoja,
                        fila,
                        fila,
                        1,
                        9
                );
            }

            CellStyle estiloTitulo =
                    ConfiguracionExcel
                            .crearEstiloTitulo(
                                    libro
                            );

            UtilExcel.establecerTexto(
                    hoja,
                    0,
                    1,
                    UtilExcel.aMayusculasSeguras(
                            safe(request.getEmpresa())
                    ),
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    1,
                    1,
                    construirTitulo(request),
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    2,
                    1,
                    "FECHA DE CORTE: "
                            + formatearFecha(
                                    request.getFechaCorte()
                            ),
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    3,
                    1,
                    "N° REGISTROS: "
                            + contarContratos(request),
                    estiloTitulo
            );

            crearEncabezadosExcel(
                    libro,
                    hoja
            );

            int filaDatosInicio =
                    FILA_ENCABEZADO_EXCEL + 1;

            int filaActual =
                    filaDatosInicio;

            AtomicInteger numeroRegistro =
                    new AtomicInteger(1);

            if (request.getGrupos() != null) {
                for (
                        GrupoTiempoServicioDTO grupo
                        : request.getGrupos()
                ) {
                    filaActual =
                            agregarGrupoExcel(
                                    hoja,
                                    filaActual,
                                    grupo,
                                    numeroRegistro
                            );
                }
            }

            int ultimaFila =
                    filaActual == filaDatosInicio
                            ? FILA_ENCABEZADO_EXCEL
                            : filaActual - 1;

            aplicarEstilosExcel(
                    libro,
                    hoja,
                    filaDatosInicio,
                    ultimaFila
            );

            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TiempoServicioTabla",
                        FILA_ENCABEZADO_EXCEL,
                        0,
                        ultimaFila,
                        ENCABEZADOS.length - 1,
                        true,
                        FILTROS_EXCEL
                );
            }

            libro.write(baos);

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;

        } catch (Exception e) {
            throw new ReportBuildException(
                    "No se pudo generar TiempoServicio.xlsx",
                    e
            );
        }
    }

    private void agregarLogoExcel(
            XSSFWorkbook libro,
            XSSFSheet hoja,
            String logoBase64
    ) {
        if (
                logoBase64 == null ||
                logoBase64.isBlank()
        ) {
            return;
        }

        byte[] logo =
                UtilExcel.decodificarImagenBase64(
                        logoBase64
                );

        if (
                logo != null &&
                logo.length > 0
        ) {
            UtilExcel.insertarLogoEstandar(
                    libro,
                    hoja,
                    logo
            );
        }
    }

    private void crearEncabezadosExcel(
            XSSFWorkbook libro,
            XSSFSheet hoja
    ) {
        Row filaEncabezado =
                UtilExcel.asegurarFila(
                        hoja,
                        FILA_ENCABEZADO_EXCEL
                );

        for (
                int columna = 0;
                columna < ENCABEZADOS.length;
                columna++
        ) {
            UtilExcel.establecerTexto(
                    filaEncabezado,
                    columna,
                    ENCABEZADOS[columna],
                    null
            );
        }

        CellStyle estiloEncabezado =
                ConfiguracionExcel
                        .crearEstiloEncabezadoTabla(
                                libro
                        );

        UtilExcel.aplicarEstiloAFila(
                filaEncabezado,
                ENCABEZADOS.length,
                estiloEncabezado
        );

        UtilExcel.establecerAnchosColumnas(
                hoja,
                ANCHOS_EXCEL
        );

        filaEncabezado.setHeightInPoints(
                20f
        );
    }

    private int agregarGrupoExcel(
            XSSFSheet hoja,
            int filaActual,
            GrupoTiempoServicioDTO grupo,
            AtomicInteger numeroRegistro
    ) {
        if (
                grupo == null ||
                grupo.getEmpleados() == null
        ) {
            return filaActual;
        }

        for (
                EmpleadoTiempoServicioDTO empleado
                : grupo.getEmpleados()
        ) {
            filaActual =
                    agregarEmpleadoExcel(
                            hoja,
                            filaActual,
                            empleado,
                            numeroRegistro
                    );
        }

        return filaActual;
    }

    private int agregarEmpleadoExcel(
            XSSFSheet hoja,
            int filaActual,
            EmpleadoTiempoServicioDTO empleado,
            AtomicInteger numeroRegistro
    ) {
        if (
                empleado == null ||
                empleado.getTiempoServicio() == null
        ) {
            return filaActual;
        }

        String nombreCompleto =
                construirNombreEmpleado(
                        empleado
                );

        for (
                ContratoTiempoServicioDTO contrato
                : empleado.getTiempoServicio()
        ) {
            if (contrato == null) {
                continue;
            }

            Row fila =
                    UtilExcel.asegurarFila(
                            hoja,
                            filaActual++
                    );

            int columna = 0;

            UtilExcel.establecerValor(
                    fila,
                    columna++,
                    numeroRegistro.getAndIncrement(),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    safe(
                            empleado.getIdentificacion()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    safe(
                            empleado.getCodigo()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    nombreCompleto,
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    firstNonEmpty(
                            contrato.getDepartamento(),
                            empleado.getDepartamento()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    firstNonEmpty(
                            contrato.getCargo(),
                            empleado.getCargo()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    firstNonEmpty(
                            contrato.getRegimen(),
                            empleado.getRegimen()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    formatearFecha(
                            contrato.getFechaIngreso()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna++,
                    formatearFechaOpcional(
                            contrato.getFechaSalida()
                    ),
                    null
            );

            UtilExcel.establecerTexto(
                    fila,
                    columna,
                    obtenerTiempoServicio(
                            contrato
                    ),
                    null
            );
        }

        return filaActual;
    }

    private void aplicarEstilosExcel(
            XSSFWorkbook libro,
            XSSFSheet hoja,
            int filaDatosInicio,
            int ultimaFila
    ) {
        if (ultimaFila < filaDatosInicio) {
            return;
        }

        CellStyle estiloCentroBorde =
                ConfiguracionExcel
                        .crearEstiloCentroConBorde(
                                libro
                        );

        CellStyle estiloIzquierdaBorde =
                ConfiguracionExcel
                        .crearEstiloIzquierdaConBorde(
                                libro
                        );

        /*
         * ITEM, IDENTIFICACIÓN y CÓDIGO centrados.
         */
        UtilExcel.aplicarEstiloARegion(
                hoja,
                filaDatosInicio,
                ultimaFila,
                0,
                2,
                estiloCentroBorde,
                true
        );

        /*
         * Nombre, departamento, cargo y régimen a la izquierda.
         */
        UtilExcel.aplicarEstiloARegion(
                hoja,
                filaDatosInicio,
                ultimaFila,
                3,
                6,
                estiloIzquierdaBorde,
                true
        );

        /*
         * Fechas y tiempo de servicio centrados.
         */
        UtilExcel.aplicarEstiloARegion(
                hoja,
                filaDatosInicio,
                ultimaFila,
                7,
                9,
                estiloCentroBorde,
                true
        );
    }

    // =====================================================================
    // VALIDACIONES
    // =====================================================================

    private void validarRequest(
            ReporteTiempoServicioRequest request
    ) {
        if (request == null) {
            throw new IllegalArgumentException(
                    "La solicitud del reporte es obligatoria."
            );
        }

        if (
                request.getGrupos() == null ||
                request.getGrupos().isEmpty()
        ) {
            throw new IllegalArgumentException(
                    "No existen grupos para generar el reporte."
            );
        }

        if (contarContratos(request) <= 0) {
            throw new IllegalArgumentException(
                    "No existen contratos para generar el reporte."
            );
        }
    }

    private int contarContratos(
            ReporteTiempoServicioRequest request
    ) {
        int total = 0;

        if (
                request == null ||
                request.getGrupos() == null
        ) {
            return total;
        }

        for (
                GrupoTiempoServicioDTO grupo
                : request.getGrupos()
        ) {
            if (
                    grupo == null ||
                    grupo.getEmpleados() == null
            ) {
                continue;
            }

            for (
                    EmpleadoTiempoServicioDTO empleado
                    : grupo.getEmpleados()
            ) {
                if (
                        empleado == null ||
                        empleado.getTiempoServicio() == null
                ) {
                    continue;
                }

                for (
                        ContratoTiempoServicioDTO contrato
                        : empleado.getTiempoServicio()
                ) {
                    if (contrato != null) {
                        total++;
                    }
                }
            }
        }

        return total;
    }

    // =====================================================================
    // HELPERS
    // =====================================================================

    private String construirTitulo(
            ReporteTiempoServicioRequest request
    ) {
        if (
                request.getTitulo() != null &&
                !request.getTitulo().isBlank()
        ) {
            return request
                    .getTitulo()
                    .trim()
                    .toUpperCase();
        }

        String estado =
                Integer.valueOf(1).equals(
                        request.getOpcionBusqueda()
                )
                        ? "USUARIOS ACTIVOS"
                        : "USUARIOS INACTIVOS";

        return NOMBRE_REPORTE
                + " - "
                + estado;
    }

    private String construirNombreEmpleado(
            EmpleadoTiempoServicioDTO empleado
    ) {
        if (empleado == null) {
            return "";
        }

        return (
                safe(empleado.getApellido())
                        + " "
                        + safe(empleado.getNombre())
        ).trim();
    }

    private String obtenerTiempoServicio(
            ContratoTiempoServicioDTO contrato
    ) {
        if (contrato == null) {
            return "";
        }

        String tiempo =
                safe(
                        contrato.getTiempoServicio()
                );

        if (!tiempo.isBlank()) {
            return tiempo;
        }

        int anios =
                contrato.getAniosServicio() == null
                        ? 0
                        : contrato.getAniosServicio();

        int meses =
                contrato.getMesesServicio() == null
                        ? 0
                        : contrato.getMesesServicio();

        int dias =
                contrato.getDiasServicio() == null
                        ? 0
                        : contrato.getDiasServicio();

        return construirTextoTiempoServicio(
                anios,
                meses,
                dias
        );
    }

    private String construirTextoTiempoServicio(
            int anios,
            int meses,
            int dias
    ) {
        String textoAnios =
                anios == 1
                        ? "1 año"
                        : anios + " años";

        String textoMeses =
                meses == 1
                        ? "1 mes"
                        : meses + " meses";

        String textoDias =
                dias == 1
                        ? "1 día"
                        : dias + " días";

        return textoAnios
                + ", "
                + textoMeses
                + ", "
                + textoDias;
    }

    /**
     * Convierte una fecha yyyy-MM-dd en:
     *
     * Lun. 17/09/2018
     */
    private String formatearFecha(
            String valor
    ) {
        String fechaTexto =
                safe(valor);

        if (fechaTexto.isBlank()) {
            return "";
        }

        try {
            String fechaISO =
                    fechaTexto.length() >= 10
                            ? fechaTexto.substring(0, 10)
                            : fechaTexto;

            LocalDate fecha =
                    LocalDate.parse(fechaISO);

            String dia =
                    fecha
                            .format(FORMATO_DIA)
                            .replace(".", "")
                            .trim();

            if (!dia.isBlank()) {
                dia =
                        Character.toUpperCase(
                                dia.charAt(0)
                        )
                                + dia.substring(1)
                                + ".";
            }

            return dia
                    + " "
                    + fecha.format(FORMATO_FECHA);

        } catch (
                DateTimeParseException |
                IndexOutOfBoundsException e
        ) {
            return fechaTexto;
        }
    }

    private String formatearFechaOpcional(
            String valor
    ) {
        String fecha =
                formatearFecha(valor);

        return fecha.isBlank()
                ? "-"
                : fecha;
    }

    private String firstNonEmpty(
            String principal,
            String respaldo
    ) {
        String valorPrincipal =
                safe(principal);

        if (!valorPrincipal.isBlank()) {
            return valorPrincipal;
        }

        return safe(respaldo);
    }

    private String safe(
            Object valor
    ) {
        if (valor == null) {
            return "";
        }

        String texto =
                String.valueOf(valor)
                        .trim();

        return "null".equalsIgnoreCase(texto)
                ? ""
                : texto;
    }
}