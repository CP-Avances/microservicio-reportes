package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionDetalleDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionEmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.PlanificacionHorarioMensualDTO;
import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.ReportePlanificacionRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;

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
        final float[] W_HORARIOS      = { 16f, 16f, 16f, 16f, 16f };
        final float[] W_NOMENCLATURA  = { 5f, 7f };
        final float[] W_CONTENEDORA   = { 60f, 40f };
        final int[]   W_ENCAB_EMPLE   = { 5, 3, 3 };
        final int     COLS_MES        = 7;
        final String  T_HEADER_HOR    = "DETALLE DE HORARIOS";
        final String  T_HEADER_DEF    = "DEFINICIONES";

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
            baos     = new ByteArrayOutputStream();
            writer   = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes)
            // Logo (opcional)
            if (request.getLogoBase64() != null) {
                Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                if (logo != null) document.add(logo);
            }

            // Títulos (empresa, título, periodo)
            Paragraph nombreEmpresa = new Paragraph(request.getEmpresa(), ReporteUtil.fuenteEncabezado());
            nombreEmpresa.setAlignment(Element.ALIGN_CENTER);
            document.add(nombreEmpresa);

            Paragraph tituloPrincipal = new Paragraph(request.getTitulo(), ReporteUtil.fuenteEncabezado());
            tituloPrincipal.setAlignment(Element.ALIGN_CENTER);
            document.add(tituloPrincipal);

            Paragraph periodo = new Paragraph(
                    "PERIODO DEL: " + request.getPeriodoInicio() + " AL " + request.getPeriodoFin(),
                    ReporteUtil.fuenteEncabezado()
            );
            periodo.setAlignment(Element.ALIGN_CENTER);
            document.add(periodo);
            document.add(Chunk.NEWLINE);

            // Colores
            Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
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

            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("HORARIO",                        colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("ENTRADA (E)",                    colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("INICIO ALIMENTACIÓN (I/A)",      colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("FIN ALIMENTACIÓN (F/A)",         colorPrincipal));
            tablaHorarios.addCell(ReporteUtil.celdaEncabezado("SALIDA (S)",                     colorPrincipal));

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
            tablaNomenclatura.addCell(ReporteUtil.celdaEncabezado("DESCRIPCIÓN",  colorSecundario));

            if (request.getNomenclatura() != null) {
                for (Map<String, String> def : request.getNomenclatura()) {
                    tablaNomenclatura.addCell(ReporteUtil.celdaNomenclatura(def.get("nombre"),       ReporteUtil.fuenteTexto()));
                    tablaNomenclatura.addCell(ReporteUtil.celdaNomenclaturaDescripcion(def.get("descripcion"), ReporteUtil.fuenteTexto()));
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

            document.add(tablaContenedora);

            // === BLOQUES POR EMPLEADO ===
            if (request.getDatos() != null) {
                for (PlanificacionEmpleadoDTO emp : request.getDatos()) {

                    // Encabezado de datos del empleado
                    PdfPTable encabezado = new PdfPTable(3);
                    encabezado.setWidths(W_ENCAB_EMPLE);

                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("EMPLEADO: " + emp.getApellido() + " " + emp.getNombre()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("C.C.: " + emp.getIdentificacion()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("COD: " + emp.getCodigo()));

                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("DEPARTAMENTO: " + emp.getDepartamento()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado("CARGO: " + emp.getCargo()));
                    encabezado.addCell(ReporteUtil.celdaInfoEmpleado(""));
                    document.add(encabezado);

                    Paragraph espacio = new Paragraph("", new Font());
                    espacio.setSpacingBefore(10f);
                    document.add(espacio);

                    // Meses (tablas de 7 columnas)
                    if (emp.getHorarios() != null) {
                        for (PlanificacionHorarioMensualDTO mes : emp.getHorarios()) {
                            PdfPTable tablaMes = new PdfPTable(COLS_MES);
                            tablaMes.setWidthPercentage(100);

                            PdfPCell celdaTituloMes = new PdfPCell(new Phrase(
                                    "AÑO: " + mes.getAnio() + " MES: " + mes.getMes(),
                                    ReporteUtil.fuenteEncabezado()
                            ));
                            celdaTituloMes.setColspan(COLS_MES);
                            celdaTituloMes.setBackgroundColor(colorSecundario);
                            celdaTituloMes.setHorizontalAlignment(Element.ALIGN_CENTER);
                            tablaMes.addCell(celdaTituloMes);

                            // cabecera de días (en bloques de 7)
                            for (int i = 1; i <= 31; i += 7) {
                                for (int j = i; j < i + 7; j++) {
                                    if (j <= 31) {
                                        PdfPCell celdaDia = new PdfPCell(
                                                new Phrase(String.format("%02d", j), ReporteUtil.fuenteEncabezado())
                                        );
                                        celdaDia.setBackgroundColor(colorPrincipal);
                                        celdaDia.setHorizontalAlignment(Element.ALIGN_CENTER);
                                        celdaDia.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                        tablaMes.addCell(celdaDia);
                                    } else {
                                        tablaMes.addCell("");
                                    }
                                }

                                // valores de los días
                                for (int j = i; j < i + 7; j++) {
                                    if (j <= 31) {
                                        Method getter = PlanificacionHorarioMensualDTO.class.getMethod("getDia" + j);
                                        String valor = (String) getter.invoke(mes);
                                        tablaMes.addCell(ReporteUtil.celdaCentro(valor != null ? valor : ""));
                                    } else {
                                        tablaMes.addCell("");
                                    }
                                }
                            }

                            document.add(tablaMes);

                            Paragraph espacioEntreEmpleados = new Paragraph("", new Font());
                            espacioEntreEmpleados.setSpacingBefore(25f);
                            document.add(espacioEntreEmpleados);
                        }
                    }
                }
            }

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
                try { document.close(); } catch (Exception ignore) {}
            }
            if (writer != null) {
                try { writer.close(); } catch (Exception ignore) {}
            }
            if (baos != null) {
                try { baos.close(); } catch (Exception ignore) {}
            }
        }
    }

        // =========================
    // XLSX (nuevo) - 3 hojas
    // =========================
    public byte[] generarReporteXLSX(ReportePlanificacionRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            // === Hoja 1: Planificacion horaria ===
            XSSFSheet hojaPlan = libro.createSheet("Planificacion horaria");

            // Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hojaPlan, logo); // A1:B5

            // Merges B1:AP5 (B=1 .. AP=41 zero-based)
            for (int row = 0; row <= 4; row++) {
                UtilExcel.combinarCeldas(hojaPlan, row, row, 1, 41);
            }

            // Títulos (empresa, título, periodo)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hojaPlan, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? "PLANIFICACION HORARIA"
                    : request.getTitulo();
            UtilExcel.establecerTexto(hojaPlan, 1, 1, UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);

            String periodo = "PERIODO DEL REPORTE: " +
                    safe(request.getPeriodoInicio()) + " AL " + safe(request.getPeriodoFin());
            UtilExcel.establecerTexto(hojaPlan, 2, 1, periodo, estiloTitulo);

            // Encabezados (fila 6 -> idx 5)
            final int filaEnc1 = 5;
            String[] encabezados1 = construirHeadersPlanificacion();
            int[] anchos1 = construirAnchosPlanificacion();

            Row header1 = UtilExcel.asegurarFila(hojaPlan, filaEnc1);
            for (int c = 0; c < encabezados1.length; c++) {
                UtilExcel.establecerTexto(header1, c, encabezados1[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(header1, encabezados1.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hojaPlan, anchos1);
            hojaPlan.getRow(filaEnc1).setHeightInPoints(18f);

            // Cuerpo
            int filaDatosIni1 = filaEnc1 + 1;
            int filaAct1 = filaDatosIni1;
            int item = 1;

            List<PlanificacionEmpleadoDTO> empleados = request.getDatos();
            if (empleados != null) {
                for (PlanificacionEmpleadoDTO emp : empleados) {
                    if (emp.getHorarios() == null) continue;

                    String apeNom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                    String codigo = safe(emp.getCodigo());
                    String ident = safe(emp.getIdentificacion());
                    String suc = safe(emp.getSucursal());
                    String cui = safe(emp.getCiudad());
                    String reg = safe(emp.getRegimen());
                    String dep = safe(emp.getDepartamento());
                    String car = safe(emp.getCargo());

                    for (PlanificacionHorarioMensualDTO h : emp.getHorarios()) {
                        Row r = UtilExcel.asegurarFila(hojaPlan, filaAct1++);
                        int col = 0;
                        UtilExcel.establecerValor(r, col++, item++, null);                 // ITEM
                        UtilExcel.establecerTexto(r, col++, codigo, null);                 // CÓDIGO
                        UtilExcel.establecerTexto(r, col++, apeNom, null);                 // NOMBRE EMPLEADO
                        UtilExcel.establecerTexto(r, col++, ident, null);                  // IDENTIFICACIÓN
                        UtilExcel.establecerTexto(r, col++, cui, null);                    // CIUDAD
                        UtilExcel.establecerTexto(r, col++, suc, null);                    // SUCURSAL
                        UtilExcel.establecerTexto(r, col++, reg, null);                    // RÉGIMEN
                        UtilExcel.establecerTexto(r, col++, dep, null);                    // DEPARTAMENTO
                        UtilExcel.establecerTexto(r, col++, car, null);                    // CARGO
                        UtilExcel.establecerTexto(r, col++, safe(h.getAnio()), null);      // AÑO
                        UtilExcel.establecerTexto(r, col++, safe(h.getMes()), null);       // MES

                        // Días 01..31 vía reflexión (getDia1..getDia31)
                        for (int d = 1; d <= 31; d++) {
                            String val = obtenerDia(h, d);
                            UtilExcel.establecerTexto(r, col++, val, null);
                        }
                    }
                }
            }

            int ultimaFila1 = (filaAct1 == filaDatosIni1) ? filaEnc1 : (filaAct1 - 1);

            // Estilos cuerpo
            CellStyle centroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle izqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado
            UtilExcel.aplicarEstiloARegion(hojaPlan, filaEnc1, filaEnc1, 0, encabezados1.length - 1, centroBorde, true);

            if (ultimaFila1 >= filaDatosIni1) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hojaPlan, filaDatosIni1, ultimaFila1, 0, 0, centroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hojaPlan, filaDatosIni1, ultimaFila1, 1, encabezados1.length - 1, izqBorde, true);
            }

            // Tabla estilizada + filtros (ITEM sin filtro)
            if (ultimaFila1 >= filaDatosIni1) {
                boolean[] filtros = new boolean[encabezados1.length];
                for (int i = 0; i < filtros.length; i++) filtros[i] = true;
                filtros[0] = false;

                UtilExcel.crearTablaEstilizada(
                        hojaPlan,
                        "ListaHorarios",
                        filaEnc1, 0,
                        ultimaFila1, encabezados1.length - 1,
                        true,
                        filtros
                );
            }

            // === Hoja 2: Detalle Horarios ===
            XSSFSheet hojaDet = libro.createSheet("Detalle Horarios");
            UtilExcel.insertarLogoEstandar(libro, hojaDet, logo); // A1:B5

            // Merges B1:F5
            for (int row = 0; row <= 4; row++) {
                UtilExcel.combinarCeldas(hojaDet, row, row, 1, 5);
            }

            UtilExcel.establecerTexto(hojaDet, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hojaDet, 1, 1, "PLANIFICACION HORARIA", estiloTitulo);
            UtilExcel.establecerTexto(hojaDet, 2, 1, periodo, estiloTitulo);

            final int filaEnc2 = 5;
            String[] encabezados2 = { "ITEM", "CÓDIGO", "ENTRADA (E)", "INICIO ALIMENTACIÓN (I/A)", "FIN ALIMENTACIÓN (F/A)", "SALIDA (S)" };
            int[] anchos2 = { 10, 20, 20, 40, 40, 20 };

            Row header2 = UtilExcel.asegurarFila(hojaDet, filaEnc2);
            for (int c = 0; c < encabezados2.length; c++) {
                UtilExcel.establecerTexto(header2, c, encabezados2[c], null);
            }
            UtilExcel.aplicarEstiloAFila(header2, encabezados2.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hojaDet, anchos2);
            hojaDet.getRow(filaEnc2).setHeightInPoints(18f);

            int filaDatosIni2 = filaEnc2 + 1;
            int filaAct2 = filaDatosIni2;
            int item2 = 1;

            List<PlanificacionDetalleDTO> detalle = request.getDetalle_acciones();
            if (detalle != null) {
                for (PlanificacionDetalleDTO d : detalle) {
                    Row r = UtilExcel.asegurarFila(hojaDet, filaAct2++);
                    UtilExcel.establecerValor (r, 0, item2++, null);
                    UtilExcel.establecerTexto (r, 1, safe(d.getHorario()), null);
                    UtilExcel.establecerTexto (r, 2, safe(d.getEntrada_()), null);
                    UtilExcel.establecerTexto (r, 3, safe(d.getInicio_comida()), null);
                    UtilExcel.establecerTexto (r, 4, safe(d.getFin_comida()), null);
                    UtilExcel.establecerTexto (r, 5, safe(d.getSalida_()), null);
                }
            }

            int ultimaFila2 = (filaAct2 == filaDatosIni2) ? filaEnc2 : (filaAct2 - 1);

            UtilExcel.aplicarEstiloARegion(hojaDet, filaEnc2, filaEnc2, 0, encabezados2.length - 1, centroBorde, true);
            if (ultimaFila2 >= filaDatosIni2) {
                UtilExcel.aplicarEstiloARegion(hojaDet, filaDatosIni2, ultimaFila2, 0, 0, centroBorde, true);
                UtilExcel.aplicarEstiloARegion(hojaDet, filaDatosIni2, ultimaFila2, 1, encabezados2.length - 1, izqBorde, true);
            }

            if (ultimaFila2 >= filaDatosIni2) {
                boolean[] filtros = new boolean[encabezados2.length];
                for (int i = 0; i < filtros.length; i++) filtros[i] = true;
                filtros[0] = false;
                UtilExcel.crearTablaEstilizada(
                        hojaDet,
                        "ListaDetaHorarios",
                        filaEnc2, 0,
                        ultimaFila2, encabezados2.length - 1,
                        true,
                        filtros
                );
            }

            // === Hoja 3: Definiciones ===
            XSSFSheet hojaDef = libro.createSheet("Definiciones");
            UtilExcel.insertarLogoEstandar(libro, hojaDef, logo); // A1:B5

            // Merges B1:C5
            for (int row = 0; row <= 4; row++) {
                UtilExcel.combinarCeldas(hojaDef, row, row, 1, 2);
            }

            UtilExcel.establecerTexto(hojaDef, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hojaDef, 1, 1, "DEFINICIONES", estiloTitulo);
            UtilExcel.establecerTexto(hojaDef, 2, 1, periodo, estiloTitulo);

            final int filaEnc3 = 5;
            String[] encabezados3 = { "ITEM", "NOMENCLATURA", "DESCRIPCIÓN" };
            int[] anchos3 = { 20, 30, 40 };

            Row header3 = UtilExcel.asegurarFila(hojaDef, filaEnc3);
            for (int c = 0; c < encabezados3.length; c++) {
                UtilExcel.establecerTexto(header3, c, encabezados3[c], null);
            }
            UtilExcel.aplicarEstiloAFila(header3, encabezados3.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hojaDef, anchos3);
            hojaDef.getRow(filaEnc3).setHeightInPoints(18f);

            int filaDatosIni3 = filaEnc3 + 1;
            int filaAct3 = filaDatosIni3;
            int item3 = 1;

            List<Map<String, String>> nomen = request.getNomenclatura();
            if (nomen != null) {
                for (Map<String, String> it : nomen) {
                    Row r = UtilExcel.asegurarFila(hojaDef, filaAct3++);
                    UtilExcel.establecerValor (r, 0, item3++, null);
                    UtilExcel.establecerTexto (r, 1, safe(it.get("nombre")), null);
                    UtilExcel.establecerTexto (r, 2, safe(it.get("descripcion")), null);
                }
            }

            int ultimaFila3 = (filaAct3 == filaDatosIni3) ? filaEnc3 : (filaAct3 - 1);

            UtilExcel.aplicarEstiloARegion(hojaDef, filaEnc3, filaEnc3, 0, encabezados3.length - 1, centroBorde, true);
            if (ultimaFila3 >= filaDatosIni3) {
                UtilExcel.aplicarEstiloARegion(hojaDef, filaDatosIni3, ultimaFila3, 0, 0,    centroBorde, true);
                UtilExcel.aplicarEstiloARegion(hojaDef, filaDatosIni3, ultimaFila3, 1, 2,    izqBorde,   true);
            }

            if (ultimaFila3 >= filaDatosIni3) {
                boolean[] filtros = new boolean[encabezados3.length];
                for (int i = 0; i < filtros.length; i++) filtros[i] = true;
                filtros[0] = false;
                UtilExcel.crearTablaEstilizada(
                        hojaDef,
                        "ListaDefinicionesHorarios",
                        filaEnc3, 0,
                        ultimaFila3, encabezados3.length - 1,
                        true,
                        filtros
                );
            }

            // ==== finalizar ====
            libro.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
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
        if (v == null) return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

}