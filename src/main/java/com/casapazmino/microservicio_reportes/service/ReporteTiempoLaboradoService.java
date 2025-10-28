package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
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
public class ReporteTiempoLaboradoService {

        public byte[] generarReporteTiempoLaboradoPDF(ReporteTiempoLaboradoRequest request) {

                // ➊ DRY: constantes locales (look & feel intacto)
                final String TITULO = "REPORTE DE TIEMPO LABORADO - "
                                + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
                final String PERIODO = "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin();

                final float[] WIDTHS_COLORES = { 3f, 1.5f, 1.5f, 2f, 2f };
                final float[] WIDTHS_TITULO = { 8f, 2f };
                final float[] WIDTHS_INFO = { 4f, 4f, 4f };
                final float[] WIDTHS_ENC_DATA = {
                                0.5f, 1.8f, // N°, FECHA
                                1.2f, 1.2f, 1.2f, 1.2f, // ENTRADA(HORARIO,TIMBRE), INICIO ALIM(HORARIO,TIMBRE)
                                1.2f, 1.2f, 1.2f, 1.2f, // FIN ALIM(HORARIO,TIMBRE), SALIDA(HORARIO,TIMBRE)
                                1.5f, 1.5f, // TIEMPO PLANIFICADO (MINUTOS, HH:MM:SS)
                                1.5f, 1.5f // TIEMPO LABORADO (MINUTOS, HH:MM:SS)
                };

                final int COLS = 14;
                final int WIDTH_PERCENT_100 = 100;
                final float SPACING_AFTER_BLOQUE = 10f;
                final float PADDING_TITULOS = 5f;

                final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
                final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
                final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();
                final Color COLOR_FT = new Color(0xEE4444);
                final Color COLOR_TIEMPO_MENOR_PLAN = new Color(0x55EE44);

                Document document = null;
                PdfWriter writer = null;
                ByteArrayOutputStream baos = null;

                try {
                        // 1) Inicialización
                        baos = new ByteArrayOutputStream();
                        document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
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

                        // Contador global de registros
                        AtomicInteger totalRegistros = new AtomicInteger();
                        if (request.getGrupos() != null) {
                                request.getGrupos().forEach(g -> {
                                        if (g.getEmpleados() != null) {
                                                g.getEmpleados().forEach(emp -> totalRegistros.addAndGet(
                                                                emp.getTLaborado() != null ? emp.getTLaborado().size()
                                                                                : 0));
                                        }
                                });
                        }

                        // Leyenda de colores
                        PdfPTable colores = new PdfPTable(5);
                        colores.setWidthPercentage(WIDTH_PERCENT_100);
                        colores.setWidths(WIDTHS_COLORES);
                        colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado(" ", COLOR_FT));
                        colores.addCell(ReporteUtil.celdaEncabezado("TIEMPO LABORADO MENOR AL PLANIFICADO",
                                        Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado(" ", COLOR_TIEMPO_MENOR_PLAN));
                        colores.setSpacingAfter(SPACING_AFTER_BLOQUE);
                        document.add(colores);

                        // Título + contador
                        PdfPTable tablaTitulo = new PdfPTable(2);
                        tablaTitulo.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaTitulo.setWidths(WIDTHS_TITULO);

                        PdfPCell celdaTitulo = new PdfPCell(
                                        new Phrase("LISTA DE EMPLEADOS", ReporteUtil.fuenteEncabezado()));
                        celdaTitulo.setBackgroundColor(COLOR_SECUNDARIO);
                        celdaTitulo.setPadding(PADDING_TITULOS);
                        celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                        tablaTitulo.addCell(celdaTitulo);

                        PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + totalRegistros.get(),
                                        ReporteUtil.fuenteEncabezado()));
                        celdaContador.setBackgroundColor(COLOR_SECUNDARIO);
                        celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
                        celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        celdaContador.setPadding(PADDING_TITULOS);
                        celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                        tablaTitulo.addCell(celdaContador);

                        tablaTitulo.setSpacingAfter(SPACING_AFTER_BLOQUE);
                        document.add(tablaTitulo);

                        // Por grupo / empleado
                        if (request.getGrupos() != null) {
                                for (GrupoTiempoDTO grupo : request.getGrupos()) {
                                        if (grupo.getEmpleados() == null)
                                                continue;

                                        for (EmpleadoTiempoDTO emp : grupo.getEmpleados()) {

                                                // Info empleado con borde exterior
                                                PdfPTable infoEmpleado = new PdfPTable(3);
                                                infoEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                                                infoEmpleado.setWidths(WIDTHS_INFO);

                                                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",
                                                                emp.getIdentificacion(), COLOR_ZEBRA));
                                                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                                                emp.getApellido() + " " + emp.getNombre(),
                                                                COLOR_ZEBRA));
                                                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(),
                                                                COLOR_ZEBRA));
                                                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:",
                                                                emp.getRegimen(), COLOR_ZEBRA));
                                                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:",
                                                                emp.getDepartamento(), COLOR_ZEBRA));
                                                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:",
                                                                emp.getCargo(), COLOR_ZEBRA));

                                                PdfPTable tablaContenedora = new PdfPTable(1);
                                                tablaContenedora.setWidthPercentage(WIDTH_PERCENT_100);
                                                PdfPCell contenedor = new PdfPCell(infoEmpleado);
                                                contenedor.setPadding(0);
                                                contenedor.setBorder(Rectangle.BOX);
                                                tablaContenedora.addCell(contenedor);
                                                tablaContenedora.setSpacingAfter(5f);
                                                document.add(tablaContenedora);

                                                // Encabezado (agrupado) en 14 columnas
                                                PdfPTable encabezado = new PdfPTable(COLS);
                                                encabezado.setWidthPercentage(WIDTH_PERCENT_100);
                                                encabezado.setWidths(WIDTHS_ENC_DATA);

                                                encabezado.addCell(ReporteUtil.crearCelda("N°",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 2, 1));
                                                encabezado.addCell(ReporteUtil.crearCelda("FECHA",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 2, 1));

                                                encabezado.addCell(ReporteUtil.crearCelda("ENTRADA",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 1, 2));
                                                encabezado.addCell(ReporteUtil.crearCelda("INICIO ALIMENTACIÓN",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 1, 2));
                                                encabezado.addCell(ReporteUtil.crearCelda("FIN ALIMENTACIÓN",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 1, 2));
                                                encabezado.addCell(ReporteUtil.crearCelda("SALIDA",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 1, 2));

                                                encabezado.addCell(ReporteUtil.crearCelda("TIEMPO PLANIFICADO",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 2, 2));
                                                encabezado.addCell(ReporteUtil.crearCelda("TIEMPO LABORADO",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO, 2, 2));

                                                // Subcolumnas
                                                for (int i = 0; i < 4; i++) {
                                                        encabezado.addCell(ReporteUtil.crearCelda("HORARIO",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        COLOR_PRIMARIO));
                                                        encabezado.addCell(ReporteUtil.crearCelda("TIMBRE",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        COLOR_SECUNDARIO));
                                                }
                                                encabezado.addCell(ReporteUtil.crearCelda("MINUTOS",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_SECUNDARIO));
                                                encabezado.addCell(ReporteUtil.crearCelda("HH:MM:SS",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_SECUNDARIO));
                                                encabezado.addCell(ReporteUtil.crearCelda("MINUTOS",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_SECUNDARIO));
                                                encabezado.addCell(ReporteUtil.crearCelda("HH:MM:SS",
                                                                ReporteUtil.fuenteEncabezado(), COLOR_SECUNDARIO));

                                                encabezado.setSpacingAfter(0f);
                                                document.add(encabezado);

                                                // Datos
                                                PdfPTable tablaData = new PdfPTable(COLS);
                                                tablaData.setWidthPercentage(WIDTH_PERCENT_100);
                                                tablaData.setWidths(WIDTHS_ENC_DATA);

                                                int contador = 1;
                                                double totalPlanificadosMin = 0d;
                                                double totalLaboradosMin = 0d;

                                                if (emp.getTLaborado() != null) {
                                                        for (RegistroTiempoDTO reg : emp.getTLaborado()) {
                                                                Color fondo = (contador % 2 == 0) ? COLOR_ZEBRA
                                                                                : Color.WHITE;

                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                String.valueOf(contador),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                ReporteUtil.formatearFechaConDia(reg
                                                                                                .getEntrada()
                                                                                                .getFecha_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));

                                                                // ENTRADA
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                extraerHora(reg.getEntrada()
                                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                                String tEntrada = formatearTimbre(reg.getEntrada()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getEntrada().getFecha_hora_timbre());
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                tEntrada, ReporteUtil.fuenteTexto(),
                                                                                getColorTimbre(tEntrada, fondo,
                                                                                                COLOR_FT)));

                                                                // INICIO ALIMENTACIÓN
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                extraerHora(reg.getInicioAlimentacion()
                                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                                String tIniAli = formatearTimbre(reg
                                                                                .getInicioAlimentacion()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getInicioAlimentacion()
                                                                                                .getFecha_hora_timbre());
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                tIniAli, ReporteUtil.fuenteTexto(),
                                                                                getColorTimbre(tIniAli, fondo,
                                                                                                COLOR_FT)));

                                                                // FIN ALIMENTACIÓN
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                extraerHora(reg.getFinAlimentacion()
                                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                                String tFinAli = formatearTimbre(reg
                                                                                .getFinAlimentacion()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getFinAlimentacion()
                                                                                                .getFecha_hora_timbre());
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                tFinAli, ReporteUtil.fuenteTexto(),
                                                                                getColorTimbre(tFinAli, fondo,
                                                                                                COLOR_FT)));

                                                                // SALIDA
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                extraerHora(reg.getSalida()
                                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                                String tSalida = formatearTimbre(
                                                                                reg.getSalida().getFecha_hora_horario(),
                                                                                reg.getSalida().getFecha_hora_timbre());
                                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                                tSalida, ReporteUtil.fuenteTexto(),
                                                                                getColorTimbre(tSalida, fondo,
                                                                                                COLOR_FT)));

                                                                // TIEMPO PLANIFICADO / LABORADO
                                                                double minPlanificado = Double
                                                                                .parseDouble(reg.getMinPlanificados());
                                                                double minLaborado = Double
                                                                                .parseDouble(reg.getMinLaborados());
                                                                Color fondoLaborado = minLaborado < minPlanificado
                                                                                ? COLOR_TIEMPO_MENOR_PLAN
                                                                                : fondo;

                                                                tablaData.addCell(ReporteUtil.celdaCentro(
                                                                                reg.getMinPlanificados(), fondo));
                                                                tablaData.addCell(ReporteUtil.celdaCentro(
                                                                                reg.getTiempoPlanificado(), fondo));
                                                                tablaData.addCell(ReporteUtil.celdaCentro(
                                                                                reg.getMinLaborados(), fondoLaborado));
                                                                tablaData.addCell(ReporteUtil.celdaCentro(
                                                                                reg.getTiempoLaborado(),
                                                                                fondoLaborado));

                                                                totalPlanificadosMin += minPlanificado;
                                                                totalLaboradosMin += minLaborado;
                                                                contador++;
                                                        }
                                                }

                                                // Fila de totales (alineada con 14 columnas)
                                                for (int i = 0; i < 9; i++) {
                                                        PdfPCell vacia = ReporteUtil.crearCelda("",
                                                                        ReporteUtil.fuenteTexto(), Color.WHITE);
                                                        vacia.setBorder(Rectangle.NO_BORDER);
                                                        tablaData.addCell(vacia);
                                                }
                                                tablaData.addCell(ReporteUtil.crearCelda("TOTAL",
                                                                ReporteUtil.fuenteTexto(), Color.WHITE));

                                                // Totales Planificado (HH:MM:SS y MINUTOS)
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                String.format("%.2f", totalPlanificadosMin).replace(",",
                                                                                "."),
                                                                ReporteUtil.fuenteTexto(), Color.WHITE));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                convertirMinutosATiempo(totalPlanificadosMin),
                                                                ReporteUtil.fuenteTexto(), Color.WHITE));

                                                // Totales Laborado (HH:MM:SS y MINUTOS)
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                String.format("%.2f", totalLaboradosMin).replace(",",
                                                                                "."),
                                                                ReporteUtil.fuenteTexto(), Color.WHITE));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                convertirMinutosATiempo(totalLaboradosMin),
                                                                ReporteUtil.fuenteTexto(), Color.WHITE));

                                                tablaData.setSpacingAfter(SPACING_AFTER_BLOQUE);
                                                document.add(tablaData);
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
                        throw new ReportBuildException("No se pudo generar ReporteTiempoLaborado.pdf", e);
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

        public byte[] generarReporteTiempoLaboradoExcel(ReporteTiempoLaboradoRequest request) {
                // =========================
                // 0) Constantes DRY locales
                // =========================
                final String NOMBRE_HOJA = "Tiempo_laborado"; // ≤ 31 chars
                final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

                // MERGES exactos (B1:V5) → (row 0..4, col 1..21)
                final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
                final int MERGE_COL_INI = 1, MERGE_COL_FIN = 21;

                final String[] HEADERS = {
                                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL",
                                "RÉGIMEN", "DEPARTAMENTO", "CARGO", "FECHA",
                                "HORARIO ENTRADA", "TIMBRE ENTRADA",
                                "HORARIO INICIO ALIMENTACIÓN", "TIMBRE INICIO ALIMENTACIÓN",
                                "HORARIO FIN ALIMENTACIÓN", "TIMBRE FIN ALIMENTACIÓN",
                                "HORARIO SALIDA", "TIMBRE SALIDA",
                                "TIEMPO PLANIFICADO", "TIEMPO PLANIFICADO MINUTOS",
                                "TIEMPO LABORADO", "TIEMPO LABORADO MINUTOS"
                };
                final int[] ANCHOS = {
                                10, 20, 20, 28, 18, 18, 18, 20, 18, 18,
                                18, 18, 22, 22, 22, 22, 18, 18, 22, 26, 22, 26
                };
                // Filtros: ITEM sin filtro; resto con filtro
                final boolean[] FILTROS = new boolean[] {
                                false, true, true, true, true, true, true, true, true, true,
                                true, true, true, true, true, true, true, true, true, true,
                                true, true
                };

                try (XSSFWorkbook libro = new XSSFWorkbook();
                                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

                        // =========================
                        // Hoja
                        // =========================
                        XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
                        hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

                        // 1) Logo estándar A1:B5
                        byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
                        if (logo != null && logo.length > 0) {
                                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
                        }

                        // 2) MERGES B1:V5
                        for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
                        }

                        // 3) TÍTULOS
                        CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
                        UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                                        estiloTitulo);

                        String ob = safe(request.getOpcionBusqueda());
                        String activosInactivos = ("1".equals(ob) || "1".equals(String.valueOf(ob))) ? "ACTIVOS"
                                        : "INACTIVOS";
                        UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TIEMPO LABORADO - " + activosInactivos,
                                        estiloTitulo);

                        String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                                        + safe(request.getFechaFin());
                        UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

                        // 4) ENCABEZADOS + ANCHOS (fila 6 → idx 5)
                        Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
                        for (int c = 0; c < HEADERS.length; c++) {
                                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
                        }
                        CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
                        UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
                        UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
                        hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

                        // 5) CUERPO (aplanado grupos → empleados → tLaborado)
                        int filaDatosIni = FILA_ENCABEZADO + 1;
                        int filaAct = filaDatosIni;
                        int item = 1;

                        if (request.getGrupos() != null) {
                                for (GrupoTiempoDTO grupo : request.getGrupos()) {
                                        if (grupo == null || grupo.getEmpleados() == null)
                                                continue;

                                        for (EmpleadoTiempoDTO emp : grupo.getEmpleados()) {
                                                if (emp == null || emp.getTLaborado() == null)
                                                        continue;

                                                String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre()))
                                                                .trim();
                                                String ciudad = firstNonEmpty(safe(emp.getCiudad()),
                                                                safe(grupo.getCiudad()));
                                                String sucursal = firstNonEmpty(safe(emp.getSucursal()),
                                                                safe(grupo.getSucursal()));

                                                for (RegistroTiempoDTO reg : emp.getTLaborado()) {
                                                        if (reg == null)
                                                                continue;

                                                        boolean esEAS = "EAS".equalsIgnoreCase(safe(reg.getTipo()));
                                                        boolean control = "true".equalsIgnoreCase(
                                                                        String.valueOf(reg.getControl()))
                                                                        || Boolean.TRUE.equals(reg.getControl());

                                                        // FECHA (desde horario de entrada)
                                                        String fecha = safe(
                                                                        () -> reg.getEntrada().getFecha_hora_horario());

                                                        // HORARIOS (HH:mm:ss)
                                                        String entradaHorario = horaDe(safe(() -> reg.getEntrada()
                                                                        .getFecha_hora_horario()));
                                                        String salidaHorario = horaDe(safe(
                                                                        () -> reg.getSalida().getFecha_hora_horario()));
                                                        String iaHorario = esEAS
                                                                        ? horaDe(safe(() -> reg.getInicioAlimentacion()
                                                                                        .getFecha_hora_horario()))
                                                                        : "";
                                                        String faHorario = esEAS ? horaDe(safe(() -> reg
                                                                        .getFinAlimentacion().getFecha_hora_horario()))
                                                                        : "";

                                                        // TIMBRES (hora || L/FD || FT/SCA según control)
                                                        String origen = safe(reg.getOrigen());
                                                        String entradaTimbre = toTimbre(
                                                                        safe(() -> reg.getEntrada()
                                                                                        .getFecha_hora_horario()),
                                                                        safe(() -> reg.getEntrada()
                                                                                        .getFecha_hora_timbre()),
                                                                        origen, control);
                                                        String salidaTimbre = toTimbre(
                                                                        safe(() -> reg.getSalida()
                                                                                        .getFecha_hora_horario()),
                                                                        safe(() -> reg.getSalida()
                                                                                        .getFecha_hora_timbre()),
                                                                        origen, control);
                                                        String iaTimbre = esEAS ? toTimbre(
                                                                        safe(() -> reg.getInicioAlimentacion()
                                                                                        .getFecha_hora_horario()),
                                                                        safe(() -> reg.getInicioAlimentacion()
                                                                                        .getFecha_hora_timbre()),
                                                                        origen, control) : "";
                                                        String faTimbre = esEAS ? toTimbre(
                                                                        safe(() -> reg.getFinAlimentacion()
                                                                                        .getFecha_hora_horario()),
                                                                        safe(() -> reg.getFinAlimentacion()
                                                                                        .getFecha_hora_timbre()),
                                                                        origen, control) : "";

                                                        // Tiempos y minutos (si !control, replica planificado)
                                                        String tiempoPlan = safe(reg.getTiempoPlanificado());
                                                        String minPlan = normalize2(safe(reg.getMinPlanificados()));
                                                        String tiempoLab = control ? safe(reg.getTiempoLaborado())
                                                                        : tiempoPlan;
                                                        String minLab = control
                                                                        ? normalize2(safe(reg.getMinLaborados()))
                                                                        : minPlan;

                                                        // === Escritura de fila ===
                                                        Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                                                        int col = 0;

                                                        UtilExcel.establecerValor(r, col++, item++, null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(emp.getIdentificacion()), null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getCodigo()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, apenom, null);
                                                        UtilExcel.establecerTexto(r, col++, ciudad, null);
                                                        UtilExcel.establecerTexto(r, col++, sucursal, null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getRegimen()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getDepartamento()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getCargo()), null);
                                                        UtilExcel.establecerTexto(r, col++, fecha, null);

                                                        UtilExcel.establecerTexto(r, col++, entradaHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, entradaTimbre, null);

                                                        UtilExcel.establecerTexto(r, col++, iaHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, iaTimbre, null);

                                                        UtilExcel.establecerTexto(r, col++, faHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, faTimbre, null);

                                                        UtilExcel.establecerTexto(r, col++, salidaHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, salidaTimbre, null);

                                                        UtilExcel.establecerTexto(r, col++, tiempoPlan, null);
                                                        UtilExcel.establecerTexto(r, col++, minPlan, null);
                                                        UtilExcel.establecerTexto(r, col++, tiempoLab, null);
                                                        UtilExcel.establecerTexto(r, col++, minLab, null);
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
                                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde,
                                                true);
                                // Resto izquierda
                                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1,
                                                estiloIzqBorde, true);
                        }

                        // 7) Tabla estilizada + filtros
                        if (ultimaFila >= filaDatosIni) {
                                UtilExcel.crearTablaEstilizada(
                                                hoja,
                                                "TiempoLaboradoTabla",
                                                FILA_ENCABEZADO, 0,
                                                ultimaFila, HEADERS.length - 1,
                                                true,
                                                FILTROS);
                        }

                        // 8) Cierre + retorno
                        libro.write(baos);
                        return baos.toByteArray();

                } catch (IllegalArgumentException e) {
                        // Validación → 400
                        throw e;
                } catch (Exception e) {
                        // Internos → 500 uniforme
                        throw new ReportBuildException("No se pudo generar TiempoLaborado.xlsx", e);
                }
        }

        private String extraerHora(String fechaHora) {
                if (fechaHora == null || !fechaHora.contains(" "))
                        return "";
                return fechaHora.split(" ")[1];
        }

        private String formatearTimbre(String horario, String timbre) {
                String horaHorario = extraerHora(horario);

                if (horaHorario == null || horaHorario.isEmpty()) {
                        return "";
                }

                if (timbre == null || timbre.trim().isEmpty()) {
                        if ("00:00:00".equals(horaHorario) || "23:59:00".equals(horaHorario)) {
                                return "L";
                        }
                        return "FT";
                }

                return extraerHora(timbre);
        }

        private Color getColorTimbre(String valor, Color porDefecto, Color colorFT) {
                if ("FT".equals(valor))
                        return colorFT;
                return porDefecto;
        }

        private String convertirMinutosATiempo(double minutos) {
                if (minutos <= 0)
                        return "00:00:00";

                int totalSegundos = (int) Math.round(minutos * 60);
                int horas = totalSegundos / 3600;
                int mins = (totalSegundos % 3600) / 60;
                int segundos = totalSegundos % 60;

                return String.format("%02d:%02d:%02d", horas, mins, segundos);
        }

        // Seguro/null-safe para String
        private String safe(Object v) {
                if (v == null)
                        return "";
                String s = String.valueOf(v).trim();
                return "null".equalsIgnoreCase(s) ? "" : s;
        }

        // Callable null-safe
        private String safe(java.util.concurrent.Callable<String> c) {
                try {
                        String s = c.call();
                        return s == null ? "" : s;
                } catch (Exception e) {
                        return "";
                }
        }

        // Devuelve "HH:mm:ss" si viene "yyyy-MM-dd HH:mm:ss"; si no, ""
        private String horaDe(String fechaHora) {
                if (fechaHora == null || fechaHora.trim().isEmpty())
                        return "";
                int idx = fechaHora.indexOf(' ');
                if (idx < 0 || idx + 1 >= fechaHora.length())
                        return "";
                return fechaHora.substring(idx + 1);
        }

        // Regla timbre (clonando la del frontend antiguo):
        // - Si existe timbre -> HH:mm:ss del timbre
        // - Si NO hay timbre:
        // - Si origen == 'L' o 'FD' -> devuelve ese código
        // - Si control == true -> "FT"
        // - Si control == false -> "SCA"
        private String toTimbre(String horario, String timbre, String origen, boolean control) {
                if (timbre != null && !timbre.trim().isEmpty()) {
                        String hh = horaDe(timbre);
                        return hh.isEmpty() ? "FT" : hh;
                }
                if ("L".equalsIgnoreCase(origen) || "FD".equalsIgnoreCase(origen)) {
                        return origen.toUpperCase();
                }
                return control ? "FT" : "SCA";
        }

        // Primera no vacía
        private String firstNonEmpty(String a, String b) {
                return (a == null || a.isBlank()) ? (b == null ? "" : b) : a;
        }

        // Normaliza "xx,yy" -> "xx.yy" y 2 decimales
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
}
