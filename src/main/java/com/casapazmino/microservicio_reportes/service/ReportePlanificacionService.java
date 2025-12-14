package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionDetalleDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionEmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionHorarioMensualDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.ReportePlanificacionRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.TextStyle;
import java.util.Locale;


import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Service
public class ReportePlanificacionService {

    public byte[] generarReportePDF(ReportePlanificacionRequest request) {

        // DRY: constantes locales (anchos, headers)
        final float[] W_HORARIOS = { 16f, 16f, 16f, 16f, 16f };
        final float[] W_NOMENCLATURA = { 5f, 7f };
        final float[] W_CONTENEDORA = { 60f, 40f };
        final int[] W_ENCAB_EMPLE = { 5, 3, 3 };
        final String T_HEADER_HOR = "DETALLE DE HORARIOS";
        final String T_HEADER_DEF = "DEFINICIONES";

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            baos = new ByteArrayOutputStream();
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (helpers existentes)
            // Logo (opcional)
            if (request.getLogoBase64() != null) {
                Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                if (logo != null)
                    document.add(logo);
            }

            // Títulos (empresa, título, periodo)
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(safe(request.getTitulo())));

            String periodo = "PERIODO DEL: " + request.getPeriodoInicio() + " AL " + request.getPeriodoFin();
            document.add(ReporteUtil.crearTituloReporte(periodo));
            document.add(Chunk.NEWLINE);

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());

            // === TABLA DE HORARIOS ===
            PdfPTable tablaHorarios = new PdfPTable(5);
            tablaHorarios.setWidthPercentage(100f);
            tablaHorarios.setHorizontalAlignment(Element.ALIGN_LEFT);
            tablaHorarios.setSpacingAfter(20f);
            tablaHorarios.setWidths(W_HORARIOS);

            PdfPCell tituloHorarios = new PdfPCell(new Phrase(T_HEADER_HOR, ReporteUtil.fuenteEncabezado()));
            tituloHorarios.setColspan(5);
            tituloHorarios.setHorizontalAlignment(Element.ALIGN_CENTER);
            tituloHorarios.setBackgroundColor(colorPrincipal);
            tablaHorarios.addCell(tituloHorarios);

            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("HORARIO", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("ENTRADA (E)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("INICIO ALIMENTACIÓN (I/A)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("FIN ALIMENTACIÓN (F/A)", colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("SALIDA (S)", colorPrincipal));

            if (request.getDetalle_acciones() != null) {
                for (PlanificacionDetalleDTO d : request.getDetalle_acciones()) {
                    tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getHorario()));
                    tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getEntrada_()));
                    tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getInicio_comida()));
                    tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getFin_comida()));
                    tablaHorarios.addCell(ReporteUtil.celdaCentro(d.getSalida_()));
                }
            }

            // === TABLA DE NOMENCLATURA ===
            PdfPTable tablaNomenclatura = new PdfPTable(2);
            tablaNomenclatura.setWidthPercentage(80f);
            tablaNomenclatura.setHorizontalAlignment(Element.ALIGN_RIGHT);
            tablaNomenclatura.setSpacingAfter(100f);
            tablaNomenclatura.setWidths(W_NOMENCLATURA);

            PdfPCell tituloNomen = new PdfPCell(new Phrase(T_HEADER_DEF, ReporteUtil.fuenteEncabezado()));
            tituloNomen.setColspan(2);
            tituloNomen.setHorizontalAlignment(Element.ALIGN_CENTER);
            tituloNomen.setBackgroundColor(colorSecundario);
            tablaNomenclatura.addCell(tituloNomen);

            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("NOMENCLATURA", colorSecundario));
            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("DESCRIPCIÓN", colorSecundario));

            if (request.getNomenclatura() != null) {
                for (Map<String, String> def : request.getNomenclatura()) {
                    tablaNomenclatura
                            .addCell(ReporteUtil.celdaNomenclatura(def.get("nombre"), ReporteUtil.fuenteTexto()));
                    tablaNomenclatura.addCell(ReporteUtil.celdaNomenclaturaDescripcion(def.get("descripcion"),
                            ReporteUtil.fuenteTexto()));
                }
            }

            // === LAYOUT CONTENEDOR (dos columnas: horarios / definiciones) ===
            PdfPTable tablaContenedora = new PdfPTable(2);
            tablaContenedora.setWidthPercentage(70f);
            tablaContenedora.setWidths(W_CONTENEDORA);
            tablaContenedora.setSpacingBefore(5f);

            PdfPCell celdaIzquierda = new PdfPCell();
            celdaIzquierda.setBorder(Rectangle.NO_BORDER);
            celdaIzquierda.setVerticalAlignment(Element.ALIGN_TOP);
            celdaIzquierda.addElement(tablaHorarios);
            tablaContenedora.addCell(celdaIzquierda);

            PdfPCell celdaDerecha = new PdfPCell();
            celdaDerecha.setBorder(Rectangle.NO_BORDER);
            celdaDerecha.setVerticalAlignment(Element.ALIGN_TOP);
            celdaDerecha.addElement(tablaNomenclatura);
            tablaContenedora.addCell(celdaDerecha);


            // === BLOQUES POR EMPLEADO ===
            if (request.getDatos() != null) {
                for (PlanificacionEmpleadoDTO emp : request.getDatos()) {

                    // Encabezado de datos del empleado
                    PdfPTable encabezado = new PdfPTable(3);
                    encabezado.setWidthPercentage(100); // ✅ MISMO ANCHO QUE CALENDARIO
                    encabezado.setWidths(W_ENCAB_EMPLE);
                    encabezado.setSpacingAfter(0f); 



                    encabezado.addCell(
                            ReporteUtil.celdaInfoEmpleado("EMPLEADO: " + emp.getApellido() + " " + emp.getNombre()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("C.C.: " + emp.getIdentificacion()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("COD: " + emp.getCodigo()));

                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("DEPARTAMENTO: " + emp.getDepartamento()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("CARGO: " + emp.getCargo()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado(""));

                    for (PdfPCell cell : encabezado.getRow(0).getCells()) {
                        if (cell != null) cell.setBorder(Rectangle.NO_BORDER);
                    }
                    for (PdfPCell cell : encabezado.getRow(1).getCells()) {
                        if (cell != null) cell.setBorder(Rectangle.NO_BORDER);
                    }

                    PdfPTable contenedorEnc = new PdfPTable(1);
                    contenedorEnc.setWidthPercentage(100);

                    PdfPCell wrap = new PdfPCell(encabezado);
                    wrap.setBorder(Rectangle.BOX);   // ✅ solo borde exterior
                    wrap.setPadding(0f);             // pegado
                    contenedorEnc.addCell(wrap);

                    document.add(contenedorEnc);

                    // Meses (tablas de 7 columnas)
                    if (emp.getHorarios() != null) {
                        for (PlanificacionHorarioMensualDTO mes : emp.getHorarios()) {

                            PdfPTable tablaMes = crearTablaMesPlanificacion(mes, colorPrincipal, colorSecundario);
                            document.add(tablaMes);

                            Paragraph espacioEntreMeses = new Paragraph("", new Font());
                            espacioEntreMeses.setSpacingBefore(25f);
                            document.add(espacioEntreMeses);
                        }
                    }

                }
            }

            document.add(tablaContenedora);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // 400 si algún helper de entrada lo lanza
        } catch (Exception e) {
            // 500 uniforme
            throw new ReportBuildException("No se pudo generar ReportePlanificacion.pdf", e);
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
    // XLSX (nuevo) - 3 hojas
    // =========================
    public byte[] generarReporteXLSX(ReportePlanificacionRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String HOJA_PLAN = "Planificacion horaria"; // ≤31
        final String HOJA_DET = "Detalle Horarios";
        final String HOJA_DEF = "Definiciones";

        final int FILA_ENC_1 = 5; // planificacion
        final int FILA_ENC_2 = 5; // detalle
        final int FILA_ENC_3 = 5; // definiciones

        // Merges B1:AP5 (row 0..4, col 1..41)
        final int M1_FIL_INI = 0, M1_FIL_FIN = 4, M1_COL_INI = 1, M1_COL_FIN = 41;

        // Merges hoja detalle: B1:F5 (row 0..4, col 1..5)
        final int M2_FIL_INI = 0, M2_FIL_FIN = 4, M2_COL_INI = 1, M2_COL_FIN = 5;

        // Merges hoja definiciones: B1:C5 (row 0..4, col 1..2)
        final int M3_FIL_INI = 0, M3_FIL_FIN = 4, M3_COL_INI = 1, M3_COL_FIN = 2;

        final String TITULO_PLAN = "PLANIFICACION HORARIA";

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            // Estilos reutilizables
            final CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            final CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            final CellStyle centroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            final CellStyle izquierdaBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Logo (una sola decodificación para reusar en 3 hojas)
            final byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());

            // =========================
            // === Hoja 1: Planificación
            // =========================
            XSSFSheet hojaPlan = libro.createSheet(HOJA_PLAN);
            hojaPlan.createFreezePane(0, FILA_ENC_1 + 1);

            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hojaPlan, logo); // A1:B5
            }

            for (int row = M1_FIL_INI; row <= M1_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hojaPlan, row, row, M1_COL_INI, M1_COL_FIN);
            }

            // Títulos (empresa, título, periodo)
            UtilExcel.establecerTexto(hojaPlan, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? TITULO_PLAN
                    : request.getTitulo();
            UtilExcel.establecerTexto(hojaPlan, 1, 1, UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);

            final String periodo = "PERIODO DEL REPORTE: " + safe(request.getPeriodoInicio()) + " AL "
                    + safe(request.getPeriodoFin());
            UtilExcel.establecerTexto(hojaPlan, 2, 1, periodo, estiloTitulo);

            // Encabezados
            final String[] HEADERS1 = construirHeadersPlanificacion();
            final int[] ANCHOS1 = construirAnchosPlanificacion();

            Row header1 = UtilExcel.asegurarFila(hojaPlan, FILA_ENC_1);
            for (int c = 0; c < HEADERS1.length; c++) {
                UtilExcel.establecerTexto(header1, c, HEADERS1[c], null);
            }
            UtilExcel.aplicarEstiloAFila(header1, HEADERS1.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hojaPlan, ANCHOS1);
            hojaPlan.getRow(FILA_ENC_1).setHeightInPoints(18f);

            // Cuerpo
            int filaDatosIni1 = FILA_ENC_1 + 1;
            int filaAct1 = filaDatosIni1;
            int item1 = 1;

            List<PlanificacionEmpleadoDTO> empleados = request.getDatos();
            if (empleados != null) {
                for (PlanificacionEmpleadoDTO emp : empleados) {
                    if (emp == null || emp.getHorarios() == null)
                        continue;

                    String apeNom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                    String codigo = safe(emp.getCodigo());
                    String ident = safe(emp.getIdentificacion());
                    String suc = safe(emp.getSucursal());
                    String cui = safe(emp.getCiudad());
                    String reg = safe(emp.getRegimen());
                    String dep = safe(emp.getDepartamento());
                    String car = safe(emp.getCargo());

                    for (PlanificacionHorarioMensualDTO h : emp.getHorarios()) {
                        if (h == null)
                            continue;
                        Row r = UtilExcel.asegurarFila(hojaPlan, filaAct1++);
                        int col = 0;
                        UtilExcel.establecerValor(r, col++, item1++, null); // ITEM
                        UtilExcel.establecerTexto(r, col++, codigo, null); // CÓDIGO
                        UtilExcel.establecerTexto(r, col++, apeNom, null); // NOMBRE EMPLEADO
                        UtilExcel.establecerTexto(r, col++, ident, null); // IDENTIFICACIÓN
                        UtilExcel.establecerTexto(r, col++, cui, null); // CIUDAD
                        UtilExcel.establecerTexto(r, col++, suc, null); // SUCURSAL
                        UtilExcel.establecerTexto(r, col++, reg, null); // RÉGIMEN
                        UtilExcel.establecerTexto(r, col++, dep, null); // DEPARTAMENTO
                        UtilExcel.establecerTexto(r, col++, car, null); // CARGO
                        UtilExcel.establecerTexto(r, col++, safe(h.getAnio()), null); // AÑO
                        UtilExcel.establecerTexto(r, col++, safe(h.getMes()), null); // MES
                        for (int d = 1; d <= 31; d++) {
                            UtilExcel.establecerTexto(r, col++, obtenerDia(h, d), null); // D01..D31
                        }
                    }
                }
            }

            int ultimaFila1 = (filaAct1 == filaDatosIni1) ? FILA_ENC_1 : (filaAct1 - 1);

            // Estilos
            UtilExcel.aplicarEstiloARegion(hojaPlan, FILA_ENC_1, FILA_ENC_1, 0, HEADERS1.length - 1, centroBorde, true);
            if (ultimaFila1 >= filaDatosIni1) {
                UtilExcel.aplicarEstiloARegion(hojaPlan, filaDatosIni1, ultimaFila1, 0, 0, centroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hojaPlan, filaDatosIni1, ultimaFila1, 1, HEADERS1.length - 1,
                        izquierdaBorde, true);
            }

            // Tabla + filtros (ITEM sin filtro)
            if (ultimaFila1 >= filaDatosIni1) {
                boolean[] filtros1 = new boolean[HEADERS1.length];
                for (int i = 0; i < filtros1.length; i++)
                    filtros1[i] = true;
                filtros1[0] = false;
                UtilExcel.crearTablaEstilizada(
                        hojaPlan,
                        "ListaHorarios",
                        FILA_ENC_1, 0,
                        ultimaFila1, HEADERS1.length - 1,
                        true,
                        filtros1);
            }

            // =========================
            // === Hoja 2: Detalle
            // =========================
            XSSFSheet hojaDet = libro.createSheet(HOJA_DET);
            hojaDet.createFreezePane(0, FILA_ENC_2 + 1);

            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hojaDet, logo);
            }

            for (int row = M2_FIL_INI; row <= M2_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hojaDet, row, row, M2_COL_INI, M2_COL_FIN);
            }

            UtilExcel.establecerTexto(hojaDet, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hojaDet, 1, 1, TITULO_PLAN, estiloTitulo);
            UtilExcel.establecerTexto(hojaDet, 2, 1, periodo, estiloTitulo);

            final String[] HEADERS2 = { "ITEM", "CÓDIGO", "ENTRADA (E)", "INICIO ALIMENTACIÓN (I/A)",
                    "FIN ALIMENTACIÓN (F/A)", "SALIDA (S)" };
            final int[] ANCHOS2 = { 10, 20, 20, 40, 40, 20 };

            Row header2 = UtilExcel.asegurarFila(hojaDet, FILA_ENC_2);
            for (int c = 0; c < HEADERS2.length; c++) {
                UtilExcel.establecerTexto(header2, c, HEADERS2[c], null);
            }
            UtilExcel.aplicarEstiloAFila(header2, HEADERS2.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hojaDet, ANCHOS2);
            hojaDet.getRow(FILA_ENC_2).setHeightInPoints(18f);

            int filaDatosIni2 = FILA_ENC_2 + 1;
            int filaAct2 = filaDatosIni2;
            int item2 = 1;

            List<PlanificacionDetalleDTO> detalle = request.getDetalle_acciones();
            if (detalle != null) {
                for (PlanificacionDetalleDTO d : detalle) {
                    if (d == null)
                        continue;
                    Row r = UtilExcel.asegurarFila(hojaDet, filaAct2++);
                    UtilExcel.establecerValor(r, 0, item2++, null);
                    UtilExcel.establecerTexto(r, 1, safe(d.getHorario()), null);
                    UtilExcel.establecerTexto(r, 2, safe(d.getEntrada_()), null);
                    UtilExcel.establecerTexto(r, 3, safe(d.getInicio_comida()), null);
                    UtilExcel.establecerTexto(r, 4, safe(d.getFin_comida()), null);
                    UtilExcel.establecerTexto(r, 5, safe(d.getSalida_()), null);
                }
            }

            int ultimaFila2 = (filaAct2 == filaDatosIni2) ? FILA_ENC_2 : (filaAct2 - 1);

            UtilExcel.aplicarEstiloARegion(hojaDet, FILA_ENC_2, FILA_ENC_2, 0, HEADERS2.length - 1, centroBorde, true);
            if (ultimaFila2 >= filaDatosIni2) {
                UtilExcel.aplicarEstiloARegion(hojaDet, filaDatosIni2, ultimaFila2, 0, 0, centroBorde, true);
                UtilExcel.aplicarEstiloARegion(hojaDet, filaDatosIni2, ultimaFila2, 1, HEADERS2.length - 1,
                        izquierdaBorde, true);
                boolean[] filtros2 = new boolean[HEADERS2.length];
                for (int i = 0; i < filtros2.length; i++)
                    filtros2[i] = true;
                filtros2[0] = false;
                UtilExcel.crearTablaEstilizada(
                        hojaDet,
                        "ListaDetaHorarios",
                        FILA_ENC_2, 0,
                        ultimaFila2, HEADERS2.length - 1,
                        true,
                        filtros2);
            }

            // =========================
            // === Hoja 3: Definiciones
            // =========================
            XSSFSheet hojaDef = libro.createSheet(HOJA_DEF);
            hojaDef.createFreezePane(0, FILA_ENC_3 + 1);

            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hojaDef, logo);
            }

            for (int row = M3_FIL_INI; row <= M3_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hojaDef, row, row, M3_COL_INI, M3_COL_FIN);
            }

            UtilExcel.establecerTexto(hojaDef, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hojaDef, 1, 1, "DEFINICIONES", estiloTitulo);
            UtilExcel.establecerTexto(hojaDef, 2, 1, periodo, estiloTitulo);

            final String[] HEADERS3 = { "ITEM", "NOMENCLATURA", "DESCRIPCIÓN" };
            final int[] ANCHOS3 = { 20, 30, 40 };

            Row header3 = UtilExcel.asegurarFila(hojaDef, FILA_ENC_3);
            for (int c = 0; c < HEADERS3.length; c++) {
                UtilExcel.establecerTexto(header3, c, HEADERS3[c], null);
            }
            UtilExcel.aplicarEstiloAFila(header3, HEADERS3.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hojaDef, ANCHOS3);
            hojaDef.getRow(FILA_ENC_3).setHeightInPoints(18f);

            int filaDatosIni3 = FILA_ENC_3 + 1;
            int filaAct3 = filaDatosIni3;
            int item3 = 1;

            List<Map<String, String>> nomen = request.getNomenclatura();
            if (nomen != null) {
                for (Map<String, String> it : nomen) {
                    if (it == null)
                        continue;
                    Row r = UtilExcel.asegurarFila(hojaDef, filaAct3++);
                    UtilExcel.establecerValor(r, 0, item3++, null);
                    UtilExcel.establecerTexto(r, 1, safe(it.get("nombre")), null);
                    UtilExcel.establecerTexto(r, 2, safe(it.get("descripcion")), null);
                }
            }

            int ultimaFila3 = (filaAct3 == filaDatosIni3) ? FILA_ENC_3 : (filaAct3 - 1);

            UtilExcel.aplicarEstiloARegion(hojaDef, FILA_ENC_3, FILA_ENC_3, 0, HEADERS3.length - 1, centroBorde, true);
            if (ultimaFila3 >= filaDatosIni3) {
                UtilExcel.aplicarEstiloARegion(hojaDef, filaDatosIni3, ultimaFila3, 0, 0, centroBorde, true);
                UtilExcel.aplicarEstiloARegion(hojaDef, filaDatosIni3, ultimaFila3, 1, 2, izquierdaBorde, true);
                boolean[] filtros3 = new boolean[HEADERS3.length];
                for (int i = 0; i < filtros3.length; i++)
                    filtros3[i] = true;
                filtros3[0] = false;
                UtilExcel.crearTablaEstilizada(
                        hojaDef,
                        "ListaDefinicionesHorarios",
                        FILA_ENC_3, 0,
                        ultimaFila3, HEADERS3.length - 1,
                        true,
                        filtros3);
            }

            // ==== finalizar ====
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Planificacion.xlsx", e); // Interno → 500
        }
    }

    // Helpers

    private static String[] construirHeadersPlanificacion() {
        String[] base = {
                "ITEM", "CÓDIGO", "NOMBRE EMPLEADO", "IDENTIFICACIÓN", "CIUDAD",
                "SUCURSAL", "REGIMEN", "DEPARTAMENTO", "CARGO", "AÑO", "MES"
        };
        String[] headers = new String[42]; // 11 base + 31 días = 42
        System.arraycopy(base, 0, headers, 0, base.length);
        int idx = base.length;
        for (int d = 1; d <= 31; d++) {
            headers[idx++] = String.format("%02d", d);
        }
        return headers;
    }

    private static int[] construirAnchosPlanificacion() {
        int[] anchos = new int[42];
        int[] primeros = { 10, 20, 30, 20, 20, 20, 20, 20, 20, 20, 20 };
        System.arraycopy(primeros, 0, anchos, 0, primeros.length);
        for (int i = primeros.length; i < anchos.length; i++) {
            anchos[i] = 20;
        }
        return anchos;
    }

    private String obtenerDia(PlanificacionHorarioMensualDTO h, int d) {
        try {
            Method getter = PlanificacionHorarioMensualDTO.class.getMethod("getDia" + d);
            String v = (String) getter.invoke(h);
            return v != null ? v : "";
        } catch (Exception ex) {
            return "";
        }
    }

    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private static final Locale LOCALE_ES_EC =
        new Locale.Builder()
                .setLanguage("es")
                .setRegion("EC")
                .build();

    private static String nombreMesES(int mes) {
        return java.time.Month.of(mes)
                .getDisplayName(TextStyle.FULL, LOCALE_ES_EC)
                .toUpperCase(LOCALE_ES_EC);
    }


    private PdfPTable crearTablaMesPlanificacion(
            PlanificacionHorarioMensualDTO mesDto,
            Color colorPrincipal,
            Color colorSecundario
    ) {
        final int COLS = 7;
        final Locale es = LOCALE_ES_EC;

        int anio = mesDto.getAnio();
        int mes = mesDto.getMes(); // 1..12

        YearMonth ym = YearMonth.of(anio, mes);
        int diasDelMes = ym.lengthOfMonth();

        // Día 1 del mes
        LocalDate first = LocalDate.of(anio, mes, 1);

        // Queremos columnas LUN..DOM (0..6)
        int startCol = (first.getDayOfWeek().getValue() + 6) % 7; // LUN=0 ... DOM=6

        // Empezamos desde el lunes de la semana donde cae el día 1
        LocalDate start = first.minusDays(startCol);

        // Número de semanas necesarias (5 o 6)
        int totalCells = startCol + diasDelMes;
        int weeks = (int) Math.ceil(totalCells / 7.0);

        PdfPTable tablaMes = new PdfPTable(COLS);
        tablaMes.setWidthPercentage(100);

        // ===== TÍTULO DEL MES (con nombre) =====
        String titulo = "AÑO: " + anio + "  MES: " + mes + " (" + nombreMesES(mes) + ")";
        PdfPCell celdaTituloMes = new PdfPCell(new Phrase(titulo, ReporteUtil.fuenteEncabezado()));
        celdaTituloMes.setColspan(COLS);
        celdaTituloMes.setBackgroundColor(colorSecundario);
        celdaTituloMes.setHorizontalAlignment(Element.ALIGN_CENTER);
        celdaTituloMes.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celdaTituloMes.setPadding(6f);
        tablaMes.addCell(celdaTituloMes);

        // ===== CUERPO: 2 FILAS POR SEMANA (Día+N°) + (Código) =====
        for (int w = 0; w < weeks; w++) {

            // ---- fila superior: "Jueves 01" ----
            for (int col = 0; col < 7; col++) {
                LocalDate date = start.plusDays(w * 7L + col);

                // Si NO pertenece al mes, celda vacía (igual que al inicio)
                if (date.getMonthValue() != mes) {
                    PdfPCell empty = new PdfPCell(new Phrase(""));
                    empty.setFixedHeight(22f);
                    empty.setHorizontalAlignment(Element.ALIGN_CENTER);
                    empty.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    tablaMes.addCell(empty);
                } else {
                    String dowFull = date.getDayOfWeek().getDisplayName(TextStyle.FULL, es); // "jueves"
                    String label = String.format("%s %02d", capitalize(dowFull), date.getDayOfMonth());

                    PdfPCell c = new PdfPCell(new Phrase(label, ReporteUtil.fuenteEncabezado()));
                    c.setBackgroundColor(colorPrincipal);
                    c.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    c.setPadding(5f);
                    c.setFixedHeight(22f);
                    tablaMes.addCell(c);
                }
            }

            // ---- fila inferior: código del día (JO1/L/...) ----
            for (int col = 0; col < 7; col++) {
                LocalDate date = start.plusDays(w * 7L + col);

                if (date.getMonthValue() != mes) {
                    PdfPCell empty = new PdfPCell(new Phrase(""));
                    empty.setFixedHeight(20f);
                    tablaMes.addCell(empty);
                } else {
                    String valor = obtenerDia(mesDto, date.getDayOfMonth());
                    PdfPCell c = new PdfPCell(new Phrase(valor != null ? valor : "", ReporteUtil.fuenteTexto()));
                    c.setHorizontalAlignment(Element.ALIGN_CENTER);
                    c.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    c.setPadding(4f);
                    c.setFixedHeight(20f);
                    tablaMes.addCell(c);
                }
            }
        }

        return tablaMes;
    }


    private static String capitalize(String s) {
        if (s == null || s.isBlank()) return "";
        s = s.trim();
        return s.substring(0, 1).toUpperCase() + s.substring(1).toLowerCase();
    }

}