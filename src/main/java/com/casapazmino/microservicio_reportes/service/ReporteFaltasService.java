package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteFaltas.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.apache.poi.ss.usermodel.Row;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteFaltasService {

    public byte[] generarReporteFaltasPDF(ReporteFaltasRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS_TITULO = { 8f, 2f };
        final float[] WIDTHS_INFO = { 4f, 4f, 4f };
        final float[] WIDTHS_FALTAS = {
                0.5f, 1.2f, 1.1f, 1.0f,
                1.5f, 1.2f, 1.2f, 1.2f
        };
        final float[] WIDTHS_RESUMEN = { 4f, 4f, 2f };

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
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            // Títulos
            String estadoUsuarios = Integer.valueOf(1).equals(request.getOpcionBusqueda())
                    ? "ACTIVOS"
                    : "INACTIVOS";

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(
                    "REPORTE DE FALTAS - USUARIOS " + estadoUsuarios
            ));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()
            ));
            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            // Contador global
            AtomicInteger totalFaltasGeneral = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados().forEach(emp -> totalFaltasGeneral.addAndGet(emp.getFaltas().size())));

            // Título global
            PdfPTable tablaTitulo = new PdfPTable(2);
            tablaTitulo.setWidthPercentage(100);
            tablaTitulo.setWidths(WIDTHS_TITULO);

            PdfPCell celda1 = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celda1.setBackgroundColor(colorSecundario);
            celda1.setPadding(5f);
            celda1.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tablaTitulo.addCell(celda1);

            PdfPCell celda2 = new PdfPCell(
                    new Phrase("Nº Registros: " + totalFaltasGeneral.get(), ReporteUtil.fuenteEncabezado()));
            celda2.setBackgroundColor(colorSecundario);
            celda2.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celda2.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celda2.setPadding(5f);
            celda2.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tablaTitulo.addCell(celda2);

            tablaTitulo.setSpacingAfter(10f);
            document.add(tablaTitulo);

            // Grupos
            for (GrupoFaltasDTO grupo : request.getGrupos()) {
                for (EmpleadoFaltasDTO emp : grupo.getEmpleados()) {
                    int contador = 1;

                    // Info empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(WIDTHS_INFO);

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                            emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado
                            .addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    tablaContenedora.setSpacingAfter(3f);
                    document.add(tablaContenedora);

                    // Tabla de faltas
                    PdfPTable tablaFaltas = new PdfPTable(8);
                    tablaFaltas.setWidthPercentage(100);
                    tablaFaltas.setWidths(WIDTHS_FALTAS);

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "N°", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "FECHA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "ESTADO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "TIPO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "DESCRIPCIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "DESDE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "HASTA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            "DETALLE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                    for (FaltaDTO falta : emp.getFaltas()) {
                        Color fondo = contador % 2 == 0 ? zebraColor : Color.WHITE;

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(contador), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                formatearFechaSegura(falta.getFecha()), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                safe(falta.getEstado_justificacion()), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                safe(falta.getTipo_justificacion()), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                safe(falta.getNombre_justificacion()), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                formatearFechaHoraJustificacion(falta.getDesde()), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                formatearFechaHoraJustificacion(falta.getHasta()), fondo));

                        tablaFaltas.addCell(ReporteUtil.celdaCentro(
                                safe(falta.getDetalle_justificacion()), fondo));

                        contador++;
                    }

                    PdfPCell celdaTotal = ReporteUtil.crearCelda(
                            "TOTAL",
                            ReporteUtil.fuenteTexto(),
                            colorSecundario
                    );
                    celdaTotal.setColspan(7);
                    celdaTotal.setHorizontalAlignment(Element.ALIGN_RIGHT);
                    tablaFaltas.addCell(celdaTotal);

                    tablaFaltas.addCell(ReporteUtil.crearCelda(
                            String.valueOf(emp.getFaltas().size()),
                            ReporteUtil.fuenteTexto(),
                            colorSecundario
                    ));

                    tablaFaltas.setSpacingAfter(10f);
                    document.add(tablaFaltas);
                }
            }

            // Resumen general final
            boolean busquedaPorEmpleado = request.getResumen() != null
                    && Boolean.TRUE.equals(request.getResumen().getBool_emp());

            if (!busquedaPorEmpleado) {
                document.add(Chunk.NEWLINE);

                PdfPTable resumen = new PdfPTable(3);
                resumen.setWidthPercentage(100);
                resumen.setWidths(WIDTHS_RESUMEN);
                resumen.setSpacingBefore(10f);

                resumen.addCell(
                        ReporteUtil.crearCelda("TOTAL GENERAL", ReporteUtil.fuenteEncabezado(), colorSecundario));
                resumen.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorSecundario));
                resumen.addCell(ReporteUtil.crearCelda(
                        String.valueOf(totalFaltasGeneral.get()), ReporteUtil.fuenteEncabezado(), colorSecundario));

                document.add(resumen);
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // 400 si algún helper lo lanza
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteFaltas.pdf", e);
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

    public byte[] generarReporteFaltasExcel(ReporteFaltasRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Faltas"; // ≤ 31 chars
        final int FILA_ENC = 5;

        // Merges B1:J5 (row 0..4, col 1..9)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 17;

        String estadoUsuarios = Integer.valueOf(1).equals(request.getOpcionBusqueda())
                ? "ACTIVOS"
                : "INACTIVOS";

        final String TITULO_REPORTE =
                "LISTA DE FALTAS - USUARIOS " + estadoUsuarios;

        final String[] HEADERS = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "GÉNERO", "CIUDAD", "NACIONALIDAD", "SUCURSAL",
                "RÉGIMEN", "DEPARTAMENTO", "CARGO", "FECHA",
                "ESTADO", "TIPO", "DESCRIPCIÓN", "DESDE", "HASTA", "DETALLE"
        };

        final int[] ANCHOS = {
                10, 20, 20, 28,
                18, 18, 20, 20,
                20, 25, 22, 20,
                20, 20, 25, 25, 25, 20
        };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            // =========================
            // 1) Inicialización de hoja
            // =========================
            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENC + 1); // mantener encabezado visible

            // 2) Logo estándar A1:B5 (si existe)
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 3) Merges B1:J5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 4) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, TITULO_REPORTE, estiloTitulo);
            String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                    + safe(request.getFechaFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            // 5) Encabezados + anchos
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENC);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENC).setHeightInPoints(18f);

            // =========================
            // 6) Cuerpo del reporte
            // =========================
            int filaDatosIni = FILA_ENC + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getGrupos() != null) {
                for (GrupoFaltasDTO grupo : request.getGrupos()) {
                    if (grupo == null || grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoFaltasDTO emp : grupo.getEmpleados()) {
                        if (emp == null || emp.getFaltas() == null)
                            continue;

                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                        String ciudad = firstNonEmpty(safe(emp.getCiudad()), safe(grupo.getCiudad()));
                        String sucursal = firstNonEmpty(safe(emp.getSucursal()), safe(grupo.getSucursal()));
                        String generoNom = safe(emp.getGeneroNombre());
                        String nacNom = safe(emp.getNacionalidadNombre());

                        for (FaltaDTO falta : emp.getFaltas()) {
                            if (falta == null)
                                continue;

                            String fechaFmt = formatearFechaSegura(falta.getFecha());
                            String estado = safe(falta.getEstado_justificacion());
                            String tipo = safe(falta.getTipo_justificacion());
                            String descripcion = safe(falta.getNombre_justificacion());
                            String desde = formatearFechaHoraJustificacion(falta.getDesde());
                            String hasta = formatearFechaHoraJustificacion(falta.getHasta());
                            String detalle = safe(falta.getDetalle_justificacion());

                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int col = 0;

                            UtilExcel.establecerValor(r, col++, item++, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCodigo()), null);
                            UtilExcel.establecerTexto(r, col++, apenom, null);
                            UtilExcel.establecerTexto(r, col++, generoNom, null);
                            UtilExcel.establecerTexto(r, col++, ciudad, null);
                            UtilExcel.establecerTexto(r, col++, nacNom, null);
                            UtilExcel.establecerTexto(r, col++, sucursal, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getRegimen()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCargo()), null);
                            UtilExcel.establecerTexto(r, col++, fechaFmt, null);
                            UtilExcel.establecerTexto(r, col++, estado, null);
                            UtilExcel.establecerTexto(r, col++, tipo, null);
                            UtilExcel.establecerTexto(r, col++, descripcion, null);
                            UtilExcel.establecerTexto(r, col++, desde, null);
                            UtilExcel.establecerTexto(r, col++, hasta, null);
                            UtilExcel.establecerTexto(r, col++, detalle, null);
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENC : (filaAct - 1);

            // =========================
            // 7) Estilos reutilizables
            // =========================
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENC, FILA_ENC, 0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosIni,
                        ultimaFila,
                        0,
                        HEADERS.length - 1,
                        estiloCentroBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosIni,
                        ultimaFila,
                        3,
                        3,
                        estiloIzqBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosIni,
                        ultimaFila,
                        9,
                        10,
                        estiloIzqBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosIni,
                        ultimaFila,
                        14,
                        14,
                        estiloIzqBorde,
                        true
                );

                UtilExcel.aplicarEstiloARegion(
                        hoja,
                        filaDatosIni,
                        ultimaFila,
                        17,
                        17,
                        estiloIzqBorde,
                        true
                );

                boolean[] filtros = new boolean[HEADERS.length];

                for (int i = 0; i < filtros.length; i++) {
                    filtros[i] = true;
                }

                filtros[0] = false;

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "FaltasReporteTabla",
                        FILA_ENC,
                        0,
                        ultimaFila,
                        HEADERS.length - 1,
                        true,
                        filtros
                );
            }

            // =========================
            // 9) Cierre + retorno
            // =========================
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación → el controller puede mapear a 400
            throw e;
        } catch (Exception e) {
            // Interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Faltas.xlsx", e);
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

    private String formatearFechaSegura(String fecha) {
        String valor = safe(fecha);
        return valor.isBlank()
                ? ""
                : ReporteUtil.formatearFechaConDia(valor);
    }

    private String formatearFechaHoraJustificacion(String valor) {
        String texto = safe(valor);
        if (texto.isBlank()) return "";

        String[] partes = texto.split("\\s+");
        String fecha = formatearFechaSegura(partes[0]);

        if (partes.length < 2) return fecha;

        String hora = partes[1].length() >= 8
                ? partes[1].substring(0, 8)
                : partes[1];

        return fecha + " " + hora;
    }

}
