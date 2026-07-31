package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteAtrasos.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.CellStyle;
import java.io.ByteArrayOutputStream;

import java.awt.Color;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAtrasosService {

    public byte[] generarReportePDF(ReporteAtrasosRequest request) {

        // DRY: anchos y colores usados varias veces
        final float[] WIDTHS_TITULO = { 8, 2 };
        final float[] WIDTHS_ENCAB_13 = {
                0.5f, 1.2f,
                1.2f, 1.2f, 1.2f, 1.3f,
                1.0f, 1.0f,
                1.0f, 0.9f,
                1.0f,
                1.0f, 0.9f
        };

        final float[] WIDTHS_DATA_13 = {
                0.5f, 1.2f,
                1.2f, 1.2f, 1.2f, 1.3f,
                1.0f, 1.0f,
                1.0f, 0.9f,
                1.0f,
                1.0f, 0.9f
        };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 30, 30, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

           document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            String usuarios = Integer.valueOf(1).equals(request.getOpcionBusqueda())
                    ? "ACTIVOS"
                    : "INACTIVOS";
            String titulo = "REPORTE DE ATRASOS - USUARIOS " + usuarios;
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Contador global
            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    g -> g.getEmpleados().forEach(e -> contadorGlobal.addAndGet(e.getAtrasos().size())));

            // Tabla título + N° registros
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(WIDTHS_TITULO);
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezadoTablaData()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezadoTablaData()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            // Iteración por grupos / empleados
            for (GrupoAtrasoDTO grupo : request.getGrupos()) {
                for (EmpleadoAtrasoDTO emp : grupo.getEmpleados()) {

                    // Info empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                            emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado
                            .addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable contenedora = new PdfPTable(1);
                    contenedora.setWidthPercentage(100);
                    PdfPCell cont = new PdfPCell(infoEmpleado);
                    cont.setPadding(0);
                    cont.setBorder(Rectangle.BOX);
                    contenedora.addCell(cont);
                    document.add(contenedora);

                    PdfPTable encabezado = new PdfPTable(13);
                    encabezado.setWidthPercentage(100);
                    encabezado.setWidths(WIDTHS_ENCAB_13);

                    // PRIMERA FILA
                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "N°",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    2,
                                    1
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "HORARIO",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    1,
                                    2
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "TIMBRE",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorSecundario,
                                    1,
                                    2
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "TIPO JUSTIFICACIÓN",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    2,
                                    1
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "DESDE",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    2,
                                    1
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "HASTA",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    2,
                                    1
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "JUSTIFICACIÓN",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    1,
                                    2
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "TOLERANCIA",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    2,
                                    1
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "ATRASO",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal,
                                    1,
                                    2
                            )
                    );

                    // SEGUNDA FILA
                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "FECHA",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "HORA",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "FECHA",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorSecundario
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "HORA",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorSecundario
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "TIEMPO",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "DECIMAL",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "TIEMPO",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal
                            )
                    );

                    encabezado.addCell(
                            ReporteUtil.crearCelda(
                                    "DECIMAL",
                                    ReporteUtil.fuenteEncabezadoTablaData(),
                                    colorPrincipal
                            )
                    );

                    encabezado.setSpacingAfter(0f);
                    document.add(encabezado);

                    // Datos (13 columnas, como en tu código original)
                    PdfPTable tablaData = new PdfPTable(13);
                    tablaData.setWidthPercentage(100);
                    tablaData.setWidths(WIDTHS_DATA_13);

                    int contador = 1;
                    long totalSegs = 0L;
                    double totalMins = 0.0;

                    for (AtrasoDTO atraso : emp.getAtrasos()) {
                        Color fondo = contador % 2 == 0
                                ? zebraColor
                                : Color.WHITE;

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        String.valueOf(contador),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        ReporteUtil.formatearFechaConDia(
                                                safe(atraso.getFechaHorario())
                                        ),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        safe(atraso.getHoraHorario()),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        ReporteUtil.formatearFechaConDia(
                                                safe(atraso.getFechaTimbre())
                                        ),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        safe(atraso.getHoraTimbre()),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );


                        // TIPO DE JUSTIFICACIÓN:
                        // - Nombre del permiso
                        // - HORAS EXTRA
                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        obtenerTipoJustificacion(atraso),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        // DESDE:
                        // Solo aplica para permisos
                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        obtenerDesdeJustificacion(atraso),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        // HASTA:
                        // Solo aplica para permisos
                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        obtenerHastaJustificacion(atraso),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        // TIEMPO REALMENTE JUSTIFICADO
                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        obtenerTiempoJustificacion(atraso),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        // MINUTOS DECIMALES REALMENTE JUSTIFICADOS
                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        obtenerDecimalJustificacion(atraso),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );



                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        safe(atraso.getTolerancia()),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        safe(atraso.getTiempoAtraso()),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        tablaData.addCell(
                                ReporteUtil.crearCelda(
                                        normalize2(atraso.getMinutosAtraso()),
                                        ReporteUtil.fuenteTablaData(),
                                        fondo
                                )
                        );

                        contador++;

                        totalSegs += convertirTiempoAtrasoASegundos(
                                atraso.getTiempoAtraso()
                        );

                        totalMins += convertirMinutos(
                                atraso.getMinutosAtraso()
                        );
                    }

                    // Totales al final
                    Color fondoTotal = new Color(230, 240, 255);

                    // Vacías hasta la columna 9
                    for (int i = 0; i < 9; i++) {
                        PdfPCell vacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        vacia.setBorder(Rectangle.NO_BORDER);
                        tablaData.addCell(vacia);
                    }

                    // "TOTAL" en columna 10
                    tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), fondoTotal));

                    // TOLERANCIA vacía en columna 11
                    tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), fondoTotal));

                    // Tiempo total formateado
                    long h = totalSegs / 3600;
                    long m = (totalSegs % 3600) / 60;
                    long s = totalSegs % 60;
                    String totalFormateado = String.format("%02d:%02d:%02d", h, m, s);

                    // Totales en ATRASO
                    tablaData.addCell(ReporteUtil.crearCelda(totalFormateado, ReporteUtil.fuenteTexto(), fondoTotal));
                    tablaData.addCell(ReporteUtil.crearCelda(
                            String.format("%.2f", totalMins).replace(",", "."),
                            ReporteUtil.fuenteTexto(), fondoTotal));

                    tablaData.setSpacingAfter(10f);
                    document.add(tablaData);
                }
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper lanza IAEx, dejamos que el controller decida si es 400
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteAtrasos.pdf", e);
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

    public byte[] generarReporteAtrasosExcel(ReporteAtrasosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Atrasos";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:U5 (B=1 .. U=20 en 0-based)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 20;

        final String[] HEADERS = {
                "ITEM",
                "IDENTIFICACIÓN",
                "CÓDIGO",
                "APELLIDO NOMBRE",
                "CIUDAD",
                "SUCURSAL",
                "RÉGIMEN",
                "DEPARTAMENTO",
                "CARGO",
                "FECHA HORARIO",
                "HORA HORARIO",
                "FECHA TIMBRE",
                "HORA TIMBRE",
                "TIPO JUSTIFICACIÓN",
                "DESDE",
                "HASTA",
                "JUSTIFICACIÓN TIEMPO",
                "JUSTIFICACIÓN DECIMAL",
                "TOLERANCIA",
                "ATRASO",
                "ATRASO MINUTOS"
        };

        final int[] ANCHOS = {
                10,
                20,
                20,
                28,
                20,
                20,
                20,
                20,
                20,
                20,
                20,
                20,
                20,
                30,
                22,
                22,
                22,
                22,
                20,
                20,
                20
        };

        // ===========================================================
        // 1) Inicialización de recursos (try-with-resources obligatorio)
        // ===========================================================
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // =====================================
            // 2) Construcción (helpers ya existentes)
            // =====================================

            // 2.1 Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2.2 Merges B1:P5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 2.3 Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);

            String activosInactivos =
                    Integer.valueOf(1).equals(request.getOpcionBusqueda())
                            ? "ACTIVOS"
                            : "INACTIVOS";

            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE ATRASOS - " + activosInactivos, estiloTitulo);

            String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                    + safe(request.getFechaFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            // 2.4 Encabezados + anchos (fila 6 → idx 5)
            Row fh = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(fh, c, HEADERS[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(fh, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 2.5 Cuerpo (aplanado grupos → empleados → atrasos)
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getGrupos() != null) {
                for (GrupoAtrasoDTO grupo : request.getGrupos()) {
                    if (grupo == null || grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoAtrasoDTO emp : grupo.getEmpleados()) {
                        if (emp == null || emp.getAtrasos() == null)
                            continue;

                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                        String ciudad = firstNonEmpty(safe(emp.getCiudad()), safe(grupo.getCiudad()));
                        String sucursal = firstNonEmpty(safe(emp.getSucursal()), safe(grupo.getSucursal()));

                        for (AtrasoDTO reg : emp.getAtrasos()) {
                            if (reg == null)
                                continue;

                            String fechaHor = safe(reg.getFechaHorario());
                            String horaHor = safe(reg.getHoraHorario());
                            String fechaTim = safe(reg.getFechaTimbre());
                            String horaTim = safe(reg.getHoraTimbre());

                            String tipoJustificacion =
                                obtenerTipoJustificacion(reg);

                            String desdeJustificacion =
                                obtenerDesdeJustificacion(reg);

                            String hastaJustificacion =
                                obtenerHastaJustificacion(reg);

                            String justificacionTiempo =
                                obtenerTiempoJustificacion(reg);

                            String justificacionDecimal =
                                obtenerDecimalJustificacion(reg);

                            String tolerancia = safe(reg.getTolerancia());
                            String atrasoFormato = safe(reg.getTiempoAtraso());
                            String atrasoMinutos = normalize2(
                                    reg.getMinutosAtraso()
                            );
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int col = 0;

                            UtilExcel.establecerValor(r, col++, item++, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCodigo()), null);
                            UtilExcel.establecerTexto(r, col++, apenom, null);
                            UtilExcel.establecerTexto(r, col++, ciudad, null);
                            UtilExcel.establecerTexto(r, col++, sucursal, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getRegimen()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCargo()), null);

                            UtilExcel.establecerTexto(r, col++, fechaHor, null);
                            UtilExcel.establecerTexto(r, col++, horaHor, null);
                            UtilExcel.establecerTexto(r, col++, fechaTim, null);
                            UtilExcel.establecerTexto(r, col++, horaTim, null);

                                UtilExcel.establecerTexto(
                                        r,
                                        col++,
                                        tipoJustificacion,
                                        null
                                );

                                UtilExcel.establecerTexto(
                                        r,
                                        col++,
                                        desdeJustificacion,
                                        null
                                );

                                UtilExcel.establecerTexto(
                                        r,
                                        col++,
                                        hastaJustificacion,
                                        null
                                );

                                UtilExcel.establecerTexto(
                                        r,
                                        col++,
                                        justificacionTiempo,
                                        null
                                );

                                UtilExcel.establecerTexto(
                                        r,
                                        col++,
                                        justificacionDecimal,
                                        null
                                );

                            UtilExcel.establecerTexto(r, col++, tolerancia, null);
                            UtilExcel.establecerTexto(r, col++, atrasoFormato, null);
                            UtilExcel.establecerTexto(r, col++, atrasoMinutos, null);
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 2.6 Estilos de cuerpo (bordes + alineación)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con bordes
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // Texto largo a la izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 3, 3, estiloIzqBorde, true); // APELLIDO
                                                                                                            // NOMBRE
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 7, 8, estiloIzqBorde, true); // DEPTO/CARGO
                // Resto centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 2, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 4, 6, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 9, 20, estiloCentroBorde, true);

                // 2.7 Tabla con filtros (ITEM sin filtro)
                boolean[] filtros = new boolean[HEADERS.length];
                for (int i = 0; i < filtros.length; i++)
                    filtros[i] = true;
                filtros[0] = false; // ITEM sin filtro

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "AtrasosReporteTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        filtros);
            }

            // ======================
            // 3) Cierre y retorno
            // ======================
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación/entrada inválida → el controller puede mapear a 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Atrasos.xlsx", e); // Interno → 500
        }
    }

    /* ===== Helpers locales ===== */
    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String firstNonEmpty(String a, String b) {
        return (a == null || a.isBlank()) ? (b == null ? "" : b) : a;
    }

    private String normalize2(String v) {
        if (v == null || v.isBlank())
            return "0.00";
        String s = v.replace(",", ".");
        try {
            double d = Double.parseDouble(s);
            return String.format(java.util.Locale.US, "%.2f", d);
        } catch (Exception e) {
            return s;
        }
    }

    /**
         * Determina si el atraso fue justificado mediante horas extra.
         *
         * Se soportan tres formas:
         * 1. Bandera enviada por el frontend.
         * 2. Código HORA_EXTRA.
         * 3. Estado JHE de la asistencia.
         */
        private boolean esJustificacionHoraExtra(AtrasoDTO atraso) {
        if (atraso == null) {
                return false;
        }

        if (Boolean.TRUE.equals(
                atraso.getEs_justificacion_hora_extra())) {
                return true;
        }

        String tipoJustificacion = safe(
                atraso.getTipo_justificacion()
        );

        if ("HORA_EXTRA".equalsIgnoreCase(tipoJustificacion)) {
                return true;
        }

        String estadoTimbre = safe(
                atraso.getEstado_timbre()
        );

        return "JHE".equalsIgnoreCase(estadoTimbre);
        }

        /**
         * Determina si la justificación proviene de un permiso.
         */
        private boolean esJustificacionPermiso(AtrasoDTO atraso) {
        if (atraso == null) {
                return false;
        }

        if (Boolean.TRUE.equals(
                atraso.getEs_justificacion_permiso())) {
                return true;
        }

        String tipoJustificacion = safe(
                atraso.getTipo_justificacion()
        );

        if ("PERMISO".equalsIgnoreCase(tipoJustificacion)) {
                return true;
        }

        /*
        * Compatibilidad con el payload anterior:
        * si existe permiso aplicado y nombre de permiso,
        * se considera justificación por permiso.
        */
        Long permisoAplicado = atraso.getPermiso_aplicado();

        return permisoAplicado != null
                && permisoAplicado > 0
                && !safe(atraso.getTipo_permiso()).isBlank();
        }

        /**
         * Devuelve el texto que se mostrará en la columna
         * TIPO JUSTIFICACIÓN.
         */
        private String obtenerTipoJustificacion(AtrasoDTO atraso) {
        if (atraso == null) {
                return "";
        }

        /*
        * El frontend ya puede enviar el texto preparado:
        * - HORAS EXTRA
        * - PERMISO MÉDICO
        */
        String texto = safe(
                atraso.getTipo_justificacion_texto()
        );

        if (!texto.isBlank()) {
                return texto;
        }

        if (esJustificacionHoraExtra(atraso)) {
                return "HORAS EXTRA";
        }

        if (esJustificacionPermiso(atraso)) {
                String tipoPermiso = safe(
                        atraso.getTipo_permiso()
                );

                if (!tipoPermiso.isBlank()) {
                return tipoPermiso;
                }

                String descripcion = safe(
                        atraso.getDescripcion_justificacion()
                );

                return descripcion.isBlank()
                        ? "PERMISO"
                        : descripcion;
        }

        /*
        * Respaldo para solicitudes con el formato anterior.
        */
        return safe(atraso.getTipo_permiso());
        }

        /**
         * Desde y hasta solamente existen para permisos.
         * En compensación por HE se muestra un guion.
         */
        private String obtenerDesdeJustificacion(AtrasoDTO atraso) {
        if (atraso == null) {
                return "";
        }

        if (esJustificacionHoraExtra(atraso)) {
                return "-";
        }

        return safe(atraso.getDesde());
        }

        private String obtenerHastaJustificacion(AtrasoDTO atraso) {
        if (atraso == null) {
                return "";
        }

        if (esJustificacionHoraExtra(atraso)) {
                return "-";
        }

        return safe(atraso.getHasta());
        }

        /**
         * Tiempo realmente aplicado a la novedad.
         *
         * Prioridad:
         * 1. Nuevo campo general de justificación.
         * 2. Tiempo del permiso aplicado.
         * 3. Duración antigua del permiso.
         */
        private String obtenerTiempoJustificacion(AtrasoDTO atraso) {
        if (atraso == null) {
                return "";
        }

        String tiempoJustificacion = safe(
                atraso.getJustificacion_tiempo()
        );

        if (!tiempoJustificacion.isBlank()) {
                return tiempoJustificacion;
        }

        String permisoAplicadoTiempo = safe(
                atraso.getPermiso_aplicado_tiempo()
        );

        if (!permisoAplicadoTiempo.isBlank()) {
                return permisoAplicadoTiempo;
        }

        return safe(
                atraso.getPermiso_tiempo()
        );
        }

        /**
         * Minutos decimales realmente aplicados a la novedad.
         */
        private String obtenerDecimalJustificacion(AtrasoDTO atraso) {
        if (atraso == null) {
                return "";
        }

        String decimalJustificacion = safe(
                atraso.getJustificacion_decimal()
        );

        if (!decimalJustificacion.isBlank()) {
                return normalize2Opcional(
                        decimalJustificacion
                );
        }

        String permisoAplicadoDecimal = safe(
                atraso.getPermiso_aplicado_decimal()
        );

        if (!permisoAplicadoDecimal.isBlank()) {
                return normalize2Opcional(
                        permisoAplicadoDecimal
                );
        }

        return normalize2Opcional(
                atraso.getPermiso_decimal()
        );
        }

    private String normalize2Opcional(String valor) {
        String texto = safe(valor);

        if (texto.isBlank()) {
            return "";
        }

        return normalize2(texto);
    }

    public static long convertirTiempoAtrasoASegundos(String tiempo) {
        if (tiempo == null || !tiempo.matches("\\d{2}:\\d{2}:\\d{2}"))
            return 0;
        try {
            String[] partes = tiempo.split(":");
            int h = Integer.parseInt(partes[0]);
            int m = Integer.parseInt(partes[1]);
            int s = Integer.parseInt(partes[2]);
            return h * 3600 + m * 60 + s;
        } catch (Exception e) {
            return 0;
        }
    }

    public static double convertirMinutos(String minutosStr) {
        if (minutosStr == null)
            return 0;
        try {
            return Double.parseDouble(minutosStr.replace(",", "."));
        } catch (Exception e) {
            return 0;
        }
    }

}
