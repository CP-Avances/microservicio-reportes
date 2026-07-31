package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Service
public class ReporteTiempoAlimentacionService {

    public byte[] generarReporteTiempoAlimentacionPDF(ReporteTiempoAlimentacionRequest request) {

        // ➊ DRY: constantes locales (look & feel intacto)
        final String TITULO = "TIEMPO DE ALIMENTACIÓN - "
                + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
        final String PERIODO = "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin();

        final float[] WIDTHS_COLORES = {
                2.4f,
                1.35f, 0.45f,
                1.9f, 0.45f,
                1.15f, 0.45f,
                1.25f, 0.45f,
                1.35f, 0.45f
        };

        final float[] WIDTHS_TITULO = { 8f, 2f };
        final float[] WIDTHS_INFO = { 4f, 4f, 4f };

        final float[] WIDTHS_TABLA = {
                0.45f,
                1.20f,
                1.10f,
                1.10f,
                1.55f,
                1.70f,
                1.70f,
                0.95f,
                0.95f,
                1.00f,
                1.10f
        };

        final String[] ENCAB_COLORES = {
                "CÓDIGO DE COLOR",
                "FALTA TIMBRE", " ",
                "EXCESO DE ALIMENTACIÓN", " ",
                "PERMISO", " ",
                "VACACIONES", " ",
                "HORAS EXTRA", " "
        };

        final String[] ENCAB_TABLA = {
                "Nº",
                "FECHA",
                "INICIO ALIMENTACIÓN",
                "FIN ALIMENTACIÓN",
                "JUSTIFICACIÓN",
                "DESDE",
                "HASTA",
                "M. ALIMENTACIÓN",
                "M. TOMADOS",
                "M. JUSTIFICADOS",
                "M. EXCESO PENDIENTE"
        };

        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_AFTER_BLOQUE = 10f;
        final float SPACING_AFTER_CONTENEDOR = 3f;
        final float PADDING_TITULOS = 5f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();
        final Color COLOR_EXCESO = new Color(0x55EE44);
        final Color COLOR_FT = new Color(0xEE4444);
        final Color COLOR_PERMISO = new Color(0xD1A15A);
        final Color COLOR_VACACION = new Color(0xFF8800);
        final Color COLOR_HORA_EXTRA = new Color(0x7EC8E3);

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
            document.add(ReporteUtil.crearTituloReporte(TITULO));
            document.add(ReporteUtil.crearTituloPeriodo(PERIODO));

            // Leyenda de colores
            PdfPTable colores = new PdfPTable(ENCAB_COLORES.length);
            colores.setWidthPercentage(WIDTH_PERCENT_100);
            colores.setWidths(WIDTHS_COLORES);

            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[0], Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[1], Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[2], COLOR_FT));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[3], Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[4], COLOR_EXCESO));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[5], Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[6], COLOR_PERMISO));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[7], Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[8], COLOR_VACACION));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[9], Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(ENCAB_COLORES[10], COLOR_HORA_EXTRA));

            colores.setSpacingAfter(SPACING_AFTER_BLOQUE);
            document.add(colores);

            // Total de registros (minas)
            AtomicInteger contadorGlobal = new AtomicInteger();
            if (request.getGrupos() != null) {
                request.getGrupos().forEach(g -> {
                    if (g.getEmpleados() != null) {
                        g.getEmpleados().forEach(e -> contadorGlobal
                                .addAndGet(e.getAlimentacion() != null ? e.getAlimentacion().size() : 0));
                    }
                });
            }

            // Título tabla + contador
            PdfPTable tablaTitulo = new PdfPTable(2);
            tablaTitulo.setWidthPercentage(WIDTH_PERCENT_100);
            tablaTitulo.setWidths(WIDTHS_TITULO);

            PdfPCell celda1 = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celda1.setBackgroundColor(COLOR_SECUNDARIO);
            celda1.setPadding(PADDING_TITULOS);
            celda1.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tablaTitulo.addCell(celda1);

            PdfPCell celda2 = new PdfPCell(
                    new Phrase("Nº Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celda2.setBackgroundColor(COLOR_SECUNDARIO);
            celda2.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celda2.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celda2.setPadding(PADDING_TITULOS);
            celda2.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tablaTitulo.addCell(celda2);

            tablaTitulo.setSpacingAfter(SPACING_AFTER_BLOQUE);
            document.add(tablaTitulo);

            if (request.getGrupos() != null) {
                for (GrupoAlimentacionDTO grupo : request.getGrupos()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoAlimentacionDTO emp : grupo.getEmpleados()) {

                        double totalJustificadoAlimentacion = 0d;
                        double totalExcesoAlimentacion = 0d;
                        int contador = 1;

                        // Encabezado info empleado con borde exterior
                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                        infoEmpleado.setWidths(WIDTHS_INFO);

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                emp.getApellido() + " " + emp.getNombre(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), COLOR_ZEBRA));
                        infoEmpleado
                                .addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), COLOR_ZEBRA));
                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), COLOR_ZEBRA));

                        PdfPTable tablaContenedora = new PdfPTable(1);
                        tablaContenedora.setWidthPercentage(WIDTH_PERCENT_100);
                        PdfPCell contenedor = new PdfPCell(infoEmpleado);
                        contenedor.setPadding(0);
                        contenedor.setBorder(Rectangle.BOX);
                        tablaContenedora.addCell(contenedor);
                        tablaContenedora.setSpacingAfter(SPACING_AFTER_CONTENEDOR);
                        document.add(tablaContenedora);

                        // Tabla principal
                        PdfPTable tablaAlimentacion = new PdfPTable(ENCAB_TABLA.length);
                        tablaAlimentacion.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaAlimentacion.setWidths(WIDTHS_TABLA);

                        for (String h : ENCAB_TABLA) {
                            tablaAlimentacion
                                    .addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        }

                        if (emp.getAlimentacion() != null) {
                            for (RegistroAlimentacionDTO registro : emp.getAlimentacion()) {
                                Color fondo = (contador % 2 == 0) ? COLOR_ZEBRA : Color.WHITE;

                                String inicio = toHoraOrCodigo(registro.getInicioAlimentacion());
                                String fin = toHoraOrCodigo(registro.getFinAlimentacion());

                                String justificacion =
                                        obtenerJustificacion(registro);

                                String desde =
                                        obtenerDesdeJustificacion(registro);

                                String hasta =
                                        obtenerHastaJustificacion(registro);

                                boolean calculoValido = !Boolean.FALSE.equals(registro.getCalculoValido());

                                String minutosPermitidos = formatNumero(registro.getMinutosPermitidos());
                                String minutosTomados = calculoValido ? formatNumero(registro.getMinutosTomados()) : "---";
                                String minutosJustificados = calculoValido
                                    ? formatNumero(
                                            registro.getMinutosJustificacionAplicados()
                                    )
                                    : "---";
                                String minutosExceso = calculoValido ? formatNumero(registro.getMinutosExceso()) : "---";

                                Color fondoInicio = obtenerColorTimbre(
                                        inicio,
                                        fondo,
                                        COLOR_FT,
                                        COLOR_PERMISO,
                                        COLOR_VACACION,
                                        COLOR_HORA_EXTRA
                                );

                                Color fondoFin = obtenerColorTimbre(
                                        fin,
                                        fondo,
                                        COLOR_FT,
                                        COLOR_PERMISO,
                                        COLOR_VACACION,
                                        COLOR_HORA_EXTRA
                                );
                                Color fondoJustificacion =
                                    obtenerColorJustificacion(
                                            registro,
                                            fondo,
                                            COLOR_PERMISO,
                                            COLOR_VACACION,
                                            COLOR_HORA_EXTRA
                                    );
                                Color fondoExceso = calculoValido
                                        && registro.getMinutosExceso() != null
                                        && registro.getMinutosExceso() > 0
                                                ? COLOR_EXCESO
                                                : fondo;

                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(String.valueOf(contador), fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(safe(registro.getFecha()), fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(inicio, fondoInicio));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(fin, fondoFin));

                                tablaAlimentacion.addCell(
                                        ReporteUtil.celdaCentro(
                                                justificacion,
                                                fondoJustificacion
                                        )
                                );


                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(desde, fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(hasta, fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(minutosPermitidos, fondo));

                                tablaAlimentacion.addCell(
                                        ReporteUtil.celdaCentro(
                                                minutosTomados,
                                                fondo
                                        )
                                );

                                tablaAlimentacion.addCell(
                                        ReporteUtil.celdaCentro(
                                                minutosJustificados,
                                                fondo
                                        )
                                );


                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(minutosExceso, fondoExceso));

                                if (calculoValido && registro.getMinutosExceso() != null) {
                                    totalExcesoAlimentacion += registro.getMinutosExceso();
                                }

                                if (
                                    calculoValido &&
                                    registro.getMinutosJustificacionAplicados() != null
                                ) {
                                    totalJustificadoAlimentacion +=
                                            registro.getMinutosJustificacionAplicados();
                                }

                                contador++;
                            }
                        }

                        // Fila de totales: 5 vacías + "TOTAL" + total exceso
                        for (int i = 0; i < 8; i++) {
                            PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                            celdaVacia.setBorder(Rectangle.NO_BORDER);
                            tablaAlimentacion.addCell(celdaVacia);
                        }

                            tablaAlimentacion.addCell(
                                    ReporteUtil.crearCelda(
                                            "TOTAL",
                                            ReporteUtil.fuenteTexto(),
                                            Color.WHITE
                                    )
                            );

                            tablaAlimentacion.addCell(
                                    ReporteUtil.crearCelda(
                                            formatNumero(totalJustificadoAlimentacion),
                                            ReporteUtil.fuenteTexto(),
                                            Color.WHITE
                                    )
                            );

                            tablaAlimentacion.addCell(
                                    ReporteUtil.crearCelda(
                                            formatNumero(totalExcesoAlimentacion),
                                            ReporteUtil.fuenteTexto(),
                                            Color.WHITE
                                    )
                            );

                        tablaAlimentacion.setSpacingAfter(SPACING_AFTER_BLOQUE);
                        document.add(tablaAlimentacion);
                    }
                }
            }

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → el controller decidirá 400 si aplica
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteTiempoAlimentacion.pdf", e);
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

    public byte[] generarReporteTiempoAlimentacionXLSX(ReporteTiempoAlimentacionRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Tiempo_Alimentacion"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

        // MERGES exactos (B1:O5) → (row 0..4, col 1..14)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 18;

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
                "FECHA",
                "INICIO ALIMENTACIÓN",
                "FIN ALIMENTACIÓN",
                "JUSTIFICACIÓN",
                "DESDE",
                "HASTA",
                "MIN. PERMITIDOS",
                "MIN. TOMADOS",
                "MIN. JUSTIFICADOS",
                "MIN. EXCESO PENDIENTE"
        };

        final int[] ANCHOS = {
                10,
                20,
                20,
                28,
                18,
                18,
                18,
                22,
                20,
                16,
                22,
                22,
                28,
                24,
                24,
                18,
                18,
                20,
                22
        };

        final boolean[] FILTROS = new boolean[HEADERS.length];
        java.util.Arrays.fill(FILTROS, true);
        FILTROS[0] = false;

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            // =========================
            // Hoja
            // =========================
            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible el encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) Merges B1:O5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);

            String ob = safe(request.getOpcionBusqueda());
            String activosInactivos = ("1".equals(ob) || "1".equals(String.valueOf(ob))) ? "ACTIVOS" : "INACTIVOS";
            UtilExcel.establecerTexto(hoja, 1, 1, "TIEMPO DE ALIMENTACIÓN - " + activosInactivos, estiloTitulo);

            String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                    + safe(request.getFechaFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupos → empleados → alimentación)
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getGrupos() != null) {
                for (GrupoAlimentacionDTO grupo : request.getGrupos()) {
                    if (grupo == null || grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoAlimentacionDTO emp : grupo.getEmpleados()) {
                        if (emp == null)
                            continue;

                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                        if (emp.getAlimentacion() == null)
                            continue;

                        for (RegistroAlimentacionDTO reg : emp.getAlimentacion()) {
                            if (reg == null)
                                continue;

                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int col = 0;

                            // === Valores (siguiendo la lógica actual) ===
                            String fecha = safe(reg.getFecha());
                            String inicioAli = toHoraOrCodigo(reg.getInicioAlimentacion());
                            String finAli = toHoraOrCodigo(reg.getFinAlimentacion());


                            String justificacion =
                                    obtenerJustificacion(reg);

                            String desde =
                                    obtenerDesdeJustificacion(reg);

                            String hasta =
                                    obtenerHastaJustificacion(reg);

                            boolean calculoValido = !Boolean.FALSE.equals(reg.getCalculoValido());

                            String minPermitidos = formatNumero(reg.getMinutosPermitidos());
                            String minTomados = calculoValido ? formatNumero(reg.getMinutosTomados()) : "---";
                            String minJustificados = calculoValido
                                ? formatNumero(
                                        reg.getMinutosJustificacionAplicados()
                                )
                                : "---";
                            String minExceso = calculoValido ? formatNumero(reg.getMinutosExceso()) : "---";
                            // === Escritura ===
                            UtilExcel.establecerValor(r, col++, item++, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCodigo()), null);
                            UtilExcel.establecerTexto(r, col++, apenom, null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCiudad()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getSucursal()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getRegimen()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, col++, safe(emp.getCargo()), null);
                            UtilExcel.establecerTexto(r, col++, fecha, null);
                            UtilExcel.establecerTexto(r, col++, inicioAli, null);
                            UtilExcel.establecerTexto(r, col++, finAli, null);
                            UtilExcel.establecerTexto(r, col++, justificacion, null);
                            UtilExcel.establecerTexto(r, col++, desde, null);
                            UtilExcel.establecerTexto(r, col++, hasta, null);
                            UtilExcel.establecerTexto(r, col++, minPermitidos, null);
                            UtilExcel.establecerTexto(r, col++, minTomados, null);
                            UtilExcel.establecerTexto(
                                    r,
                                    col++,
                                    minJustificados,
                                    null
                            );
                            UtilExcel.establecerTexto(r, col++, minExceso, null);
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Estilos de cuerpo (por región)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // Resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde,
                        true);
            }

            // 7) Tabla estilizada + filtros
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TiempoAlimentacionTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        FILTROS);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación → 400 (lo mapea el controller)
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar TiempoAlimentacion.xlsx", e);
        }
    }

    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String toHoraOrCodigo(String valor) {
        String texto = safe(valor);
        if (texto.isEmpty()) return "FT";

        String codigo = texto.toUpperCase();
        if (
            "FT".equals(codigo) ||
            "P".equals(codigo) ||
            "V".equals(codigo) ||
            "JHE".equals(codigo)
        ) {
            return codigo;
        }

        int indiceEspacio = texto.indexOf(' ');
        int indiceT = texto.indexOf('T');
        int indice = indiceEspacio >= 0 ? indiceEspacio : indiceT;

        if (indice >= 0 && indice + 1 < texto.length()) return texto.substring(indice + 1);
        return texto;
    }

    private String obtenerJustificacion(
            RegistroAlimentacionDTO registro
    ) {
        if (registro == null) {
            return "";
        }

        String texto = safe(
                registro.getTipoJustificacionTexto()
        );

        if (!texto.isEmpty()) {
            return texto;
        }

        String tipo = safe(
                registro.getTipoJustificacion()
        ).toUpperCase();

        if ("JHE".equals(tipo)) {
            return "HORAS EXTRA";
        }

        if ("V".equals(tipo)) {
            return "VACACIONES";
        }

        if ("P".equals(tipo)) {
            String tipoPermiso = safe(
                    registro.getTipoPermiso()
            );

            return tipoPermiso.isEmpty()
                    ? "PERMISO"
                    : tipoPermiso;
        }

        return safe(registro.getTipoPermiso());
    }

    private Color obtenerColorTimbre(
            String valor,
            Color fondoNormal,
            Color colorFt,
            Color colorPermiso,
            Color colorVacacion,
            Color colorHoraExtra
    ) {
        String codigo = safe(valor).toUpperCase();

        if ("FT".equals(codigo)) {
            return colorFt;
        }

        if ("P".equals(codigo)) {
            return colorPermiso;
        }

        if ("V".equals(codigo)) {
            return colorVacacion;
        }

        if ("JHE".equals(codigo)) {
            return colorHoraExtra;
        }

        return fondoNormal;
    }
        
    
    private String formatNumero(Double valor) {
        if (valor == null || !Double.isFinite(valor)) return "0";
        if (Math.rint(valor) == valor) return String.valueOf(valor.longValue());
        return String.format("%.2f", valor).replace(",", ".");
    }

    private String formatNumero(double valor) {
        if (!Double.isFinite(valor)) {
            return "0";
        }

        if (Math.rint(valor) == valor) {
            return String.valueOf((long) valor);
        }

        return String.format(
                java.util.Locale.US,
                "%.2f",
                valor
        );
    }

    private boolean esJustificacionHoraExtra(
            RegistroAlimentacionDTO registro
    ) {
        if (registro == null) {
            return false;
        }

        if (Boolean.TRUE.equals(
                registro.getEsJustificacionHoraExtra()
        )) {
            return true;
        }

        return "JHE".equalsIgnoreCase(
                safe(registro.getTipoJustificacion())
        );
    }

    private String obtenerDesdeJustificacion(
            RegistroAlimentacionDTO registro
    ) {
        if (esJustificacionHoraExtra(registro)) {
            return "-";
        }

        return safe(registro.getDesde());
    }

    private String obtenerHastaJustificacion(
            RegistroAlimentacionDTO registro
    ) {
        if (esJustificacionHoraExtra(registro)) {
            return "-";
        }

        return safe(registro.getHasta());
    }


    private Color obtenerColorJustificacion(
            RegistroAlimentacionDTO registro,
            Color fondoNormal,
            Color colorPermiso,
            Color colorVacacion,
            Color colorHoraExtra
    ) {
        String tipo = safe(
                registro != null
                        ? registro.getTipoJustificacion()
                        : null
        ).toUpperCase();

        if ("P".equals(tipo)) {
            return colorPermiso;
        }

        if ("V".equals(tipo)) {
            return colorVacacion;
        }

        if ("JHE".equals(tipo)) {
            return colorHoraExtra;
        }

        return fondoNormal;
    }

}
