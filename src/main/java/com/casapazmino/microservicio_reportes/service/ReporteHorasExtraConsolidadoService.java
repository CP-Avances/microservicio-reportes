package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteHorasExtraConsolidadoService {

        public byte[] generarReportePDF(ReporteHorasExtraConsolidadoRequest request) {
                final float[] WIDTHS_21 = {
                                0.7f, 1.6f,
                                1.5f, 1.4f, 1.5f, 1.4f,
                                1.5f, 1.4f, 1.5f, 1.4f,
                                1.3f, 1.3f, 1.5f, 1.5f,
                                1.8f,
                                1.4f, 1.3f, 1.8f,
                                1.8f, 2.0f, 2.0f
                };

                final Color COLOR_FALTA_TIMBRE = new Color(0xEE4444);
                final Color COLOR_ATRASO = new Color(0xEEE344);
                final Color COLOR_SALIDA_ANTICIPADA = new Color(0x4499EE);
                final Color COLOR_EXCESO_ALIMENTACION = new Color(0x55EE44);

                Document document = null;
                PdfWriter writer = null;
                ByteArrayOutputStream baos = null;

                try {
                        baos = new ByteArrayOutputStream();
                        document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
                        writer = PdfWriter.getInstance(document, baos);
                        writer.setPageEvent(new ConfiguracionPaginaPDF(
                                        request.getUsuario(),
                                        request.getFraseMarcaAgua(),
                                        request.getColorPrincipal()));
                        document.open();

                        Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                        if (logo != null) {
                                document.add(logo);
                        }

                        document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
                        String titulo = "HORAS EXTRA - CONSOLIDADO - "
                                        + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
                        document.add(ReporteUtil.crearTituloReporte(titulo));
                        document.add(ReporteUtil.crearTituloPeriodo(
                                        "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

                        Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
                        Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
                        Color zebraColor = ReporteUtil.colorZebraClaro();

                        AtomicInteger totalRegistros = new AtomicInteger();
                        if (request.getGrupos() != null) {
                                request.getGrupos().forEach(g -> {
                                        if (g.getEmpleados() != null) {
                                                g.getEmpleados().forEach(emp -> {
                                                        if (emp.getTLaborado() != null) {
                                                                emp.getTLaborado().forEach(reg -> {
                                                                        List<DetalleHoraExtraDTO> detalles = safeDetalles(
                                                                                        reg.getDetalleHorasExtra());
                                                                        totalRegistros.addAndGet(detalles.isEmpty() ? 1
                                                                                        : detalles.size());
                                                                });
                                                        }
                                                });
                                        }
                                });
                        }

                        PdfPTable tituloTabla = new PdfPTable(2);
                        tituloTabla.setWidthPercentage(100);
                        tituloTabla.setWidths(new float[] { 8, 2 });
                        tituloTabla.setSpacingAfter(5f);

                        PdfPCell celdaTitulo = new PdfPCell(
                                        new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
                        celdaTitulo.setBackgroundColor(colorSecundario);
                        celdaTitulo.setPadding(5f);
                        celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                        tituloTabla.addCell(celdaTitulo);

                        PdfPCell celdaContador = new PdfPCell(
                                        new Phrase("N° Registros: " + totalRegistros.get(),
                                                        ReporteUtil.fuenteEncabezado()));
                        celdaContador.setBackgroundColor(colorSecundario);
                        celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
                        celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        celdaContador.setPadding(5f);
                        celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                        tituloTabla.addCell(celdaContador);

                        tituloTabla.setSpacingAfter(10f);
                        document.add(tituloTabla);

                        for (GrupoHorasExtraConsolidadoDT grupo : safeGrupos(request.getGrupos())) {
                                for (EmpleadoHorasExtraConsolidadoDTO emp : safeEmpleados(grupo.getEmpleados())) {

                                        double totalAtrasos = 0;
                                        double totalSalidasAnticipadas = 0;
                                        double totalAlimentacionTomado = 0;
                                        double totalAlimentacionAsignado = 0;
                                        double totalLaborado = 0;
                                        double totalHorasExtra = 0;

                                        PdfPTable infoEmpleado = new PdfPTable(3);
                                        infoEmpleado.setWidthPercentage(100);
                                        infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",
                                                        safe(emp.getIdentificacion()), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                                        (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim(),
                                                        zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:",
                                                        safe(emp.getCodigo()), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:",
                                                        safe(emp.getRegimen()), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:",
                                                        safe(emp.getDepartamento()), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:",
                                                        safe(emp.getCargo()), zebraColor));

                                        PdfPTable tablaContenedora = new PdfPTable(1);
                                        tablaContenedora.setWidthPercentage(100);
                                        PdfPCell contenedor = new PdfPCell(infoEmpleado);
                                        contenedor.setPadding(0);
                                        contenedor.setBorder(Rectangle.BOX);
                                        tablaContenedora.addCell(contenedor);
                                        tablaContenedora.setSpacingAfter(5f);
                                        document.add(tablaContenedora);

                                        PdfPTable encabezado = new PdfPTable(21);
                                        encabezado.setWidthPercentage(100);
                                        encabezado.setWidths(WIDTHS_21);

                                        encabezado.addCell(ReporteUtil.crearCelda("N°",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("FECHA",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                                        encabezado.addCell(ReporteUtil.crearCelda("ENTRADA",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("INICIO ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("FIN ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("SALIDA",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));

                                        encabezado.addCell(ReporteUtil.crearCelda("ATRASO",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("SALIDA ANTICIPADA",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("T. ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("TIEMPO LABORADO",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                                        encabezado.addCell(ReporteUtil.crearCelda("HORAS EXTRA",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("MINUTOS",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("TIPO %",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("PORCENTAJE",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("TIPO RECARGO",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCelda("OBSERVACIONES",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal, 2, 1));

                                        for (int i = 0; i < 4; i++) {
                                                encabezado.addCell(ReporteUtil.crearCelda("HORARIO",
                                                                ReporteUtil.fuenteEncabezado(), colorPrincipal));
                                                encabezado.addCell(ReporteUtil.crearCelda("TIMBRE",
                                                                ReporteUtil.fuenteEncabezado(), colorSecundario));
                                        }
                                        encabezado.addCell(ReporteUtil.crearCelda("ASIGNADO",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCelda("TOMADO",
                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal));

                                        encabezado.setSpacingAfter(0f);
                                        document.add(encabezado);

                                        PdfPTable tablaData = new PdfPTable(21);
                                        tablaData.setWidthPercentage(100);
                                        tablaData.setWidths(WIDTHS_21);

                                        int contador = 1;

                                        for (RegistroHorasExtraConsolidadoDTO reg : safeRegistros(emp.getTLaborado())) {
                                                List<DetalleHoraExtraDTO> detalles = safeDetalles(
                                                                reg.getDetalleHorasExtra());

                                                if (detalles.isEmpty()) {
                                                        detalles = new ArrayList<>();
                                                        detalles.add(new DetalleHoraExtraDTO());
                                                }

                                                for (DetalleHoraExtraDTO detalle : detalles) {
                                                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        String.valueOf(contador),
                                                                        ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        ReporteUtil.formatearFechaConDia(
                                                                                        safe(reg.getEntrada() != null
                                                                                                        ? reg.getEntrada()
                                                                                                                        .getFecha_horario()
                                                                                                        : "")),
                                                                        ReporteUtil.fuenteTexto(), fondo));

                                                        String entradaHorario = extraerHora(
                                                                        reg.getEntrada() != null ? reg.getEntrada()
                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String entradaTimbre = formatearTimbre(
                                                                        reg.getEntrada() != null ? reg.getEntrada()
                                                                                        .getFecha_hora_horario() : null,
                                                                        reg.getEntrada() != null ? reg.getEntrada()
                                                                                        .getFecha_hora_timbre() : null);
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        entradaHorario, ReporteUtil.fuenteTexto(),
                                                                        fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        entradaTimbre, ReporteUtil.fuenteTexto(),
                                                                        getColorTimbre(entradaTimbre, fondo,
                                                                                        COLOR_FALTA_TIMBRE)));

                                                        String iaHorario = extraerHora(
                                                                        reg.getInicioAlimentacion() != null ? reg
                                                                                        .getInicioAlimentacion()
                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String iaTimbre = formatearTimbre(
                                                                        reg.getInicioAlimentacion() != null ? reg
                                                                                        .getInicioAlimentacion()
                                                                                        .getFecha_hora_horario() : null,
                                                                        reg.getInicioAlimentacion() != null ? reg
                                                                                        .getInicioAlimentacion()
                                                                                        .getFecha_hora_timbre() : null);
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        iaHorario, ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        iaTimbre, ReporteUtil.fuenteTexto(),
                                                                        getColorTimbre(iaTimbre, fondo,
                                                                                        COLOR_FALTA_TIMBRE)));

                                                        String faHorario = extraerHora(
                                                                        reg.getFinAlimentacion() != null ? reg
                                                                                        .getFinAlimentacion()
                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String faTimbre = formatearTimbre(
                                                                        reg.getFinAlimentacion() != null ? reg
                                                                                        .getFinAlimentacion()
                                                                                        .getFecha_hora_horario() : null,
                                                                        reg.getFinAlimentacion() != null ? reg
                                                                                        .getFinAlimentacion()
                                                                                        .getFecha_hora_timbre() : null);
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        faHorario, ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        faTimbre, ReporteUtil.fuenteTexto(),
                                                                        getColorTimbre(faTimbre, fondo,
                                                                                        COLOR_FALTA_TIMBRE)));

                                                        String salidaHorario = extraerHora(
                                                                        reg.getSalida() != null ? reg.getSalida()
                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String salidaTimbre = formatearTimbre(
                                                                        reg.getSalida() != null ? reg.getSalida()
                                                                                        .getFecha_hora_horario() : null,
                                                                        reg.getSalida() != null ? reg.getSalida()
                                                                                        .getFecha_hora_timbre() : null);
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        salidaHorario, ReporteUtil.fuenteTexto(),
                                                                        fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        salidaTimbre, ReporteUtil.fuenteTexto(),
                                                                        getColorTimbre(salidaTimbre, fondo,
                                                                                        COLOR_FALTA_TIMBRE)));

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        convertirMinutosATiempo(reg.getMinAtrasos()),
                                                                        ReporteUtil.fuenteTexto(),
                                                                        reg.getMinAtrasos() != null
                                                                                        && reg.getMinAtrasos() > 0
                                                                                                        ? COLOR_ATRASO
                                                                                                        : fondo));

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        convertirMinutosATiempo(
                                                                                        reg.getMinSalidasAnticipadas()),
                                                                        ReporteUtil.fuenteTexto(),
                                                                        reg.getMinSalidasAnticipadas() != null && reg
                                                                                        .getMinSalidasAnticipadas() > 0
                                                                                                        ? COLOR_SALIDA_ANTICIPADA
                                                                                                        : fondo));

                                                        Double minAsignado = 0d;
                                                        if (reg.getInicioAlimentacion() != null && reg
                                                                        .getInicioAlimentacion()
                                                                        .getMinutos_alimentacion() != null) {
                                                                minAsignado = reg.getInicioAlimentacion()
                                                                                .getMinutos_alimentacion();
                                                        }

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        convertirMinutosATiempo(minAsignado),
                                                                        ReporteUtil.fuenteTexto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        convertirMinutosATiempo(
                                                                                        reg.getMinAlimentacion()),
                                                                        ReporteUtil.fuenteTexto(),
                                                                        (minAsignado != null && reg
                                                                                        .getMinAlimentacion() != null
                                                                                        && reg.getMinAlimentacion() > minAsignado)
                                                                                                        ? COLOR_EXCESO_ALIMENTACION
                                                                                                        : fondo));

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        convertirMinutosATiempo(reg.getMinLaborados()),
                                                                        ReporteUtil.fuenteTexto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        safe(detalle.getHorasExtra()),
                                                                        ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        formatDouble(detalle.getMinutosHorasExtra()),
                                                                        ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        safe(detalle.getTipoPorcentaje()),
                                                                        ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        safe(detalle.getPorcentaje()),
                                                                        ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        safe(detalle.getTipoRecargo()),
                                                                        ReporteUtil.fuenteTexto(), fondo));
                                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                                        "", ReporteUtil.fuenteTexto(), fondo));

                                                        totalAtrasos += nz(reg.getMinAtrasos());
                                                        totalSalidasAnticipadas += nz(reg.getMinSalidasAnticipadas());
                                                        totalAlimentacionTomado += nz(reg.getMinAlimentacion());
                                                        totalAlimentacionAsignado += nz(minAsignado);
                                                        totalLaborado += nz(reg.getMinLaborados());
                                                        totalHorasExtra += nz(detalle.getMinutosHorasExtra());

                                                        contador++;
                                                }
                                        }

                                        for (int i = 0; i < 14; i++) {
                                                PdfPCell vacia = ReporteUtil.crearCelda("",
                                                                ReporteUtil.fuenteTexto(), Color.WHITE);
                                                vacia.setBorder(Rectangle.NO_BORDER);
                                                tablaData.addCell(vacia);
                                        }

                                        tablaData.addCell(ReporteUtil.crearCelda("TOTAL",
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        convertirMinutosATiempo(totalHorasExtra),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        formatDouble(totalHorasExtra),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(),
                                                        Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(),
                                                        Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(),
                                                        Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(),
                                                        Color.WHITE));

                                        tablaData.setSpacingAfter(10f);
                                        document.add(tablaData);
                                }
                        }

                        document.close();
                        return baos.toByteArray();

                } catch (IllegalArgumentException e) {
                        throw e;
                } catch (Exception e) {
                        throw new ReportBuildException("No se pudo generar HorasExtraConsolidado.pdf", e);
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

        public byte[] generarReporteXLSX(ReporteHorasExtraConsolidadoRequest request) {
                final String NOMBRE_HOJA = "HorasExtra_Consolidado";
                final int FILA_ENCABEZADO = 5;

                final String[] HEADERS = {
                                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL", "RÉGIMEN",
                                "DEPARTAMENTO", "CARGO", "FECHA",
                                "HORARIO ENTRADA", "TIMBRE ENTRADA",
                                "HORARIO INICIO ALIMENTACIÓN", "TIMBRE INICIO ALIMENTACIÓN",
                                "HORARIO FIN ALIMENTACIÓN", "TIMBRE FIN ALIMENTACIÓN",
                                "HORARIO SALIDA", "TIMBRE SALIDA",
                                "ATRASO", "SALIDA ANTICIPADA",
                                "TIEMPO ALIMENTACIÓN ASIGNADO", "TIEMPO ALIMENTACIÓN",
                                "TIEMPO LABORADO",
                                "HORAS EXTRA", "MINUTOS", "TIPO %", "PORCENTAJE", "TIPO RECARGO"
                };

                final int[] ANCHOS = {
                                10, 18, 14, 24, 18, 18, 18,
                                20, 18, 18,
                                18, 18,
                                22, 22,
                                22, 22,
                                18, 18,
                                16, 18,
                                22, 22,
                                18,
                                18, 14, 20, 16, 22
                };

                try (XSSFWorkbook libro = new XSSFWorkbook();
                                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

                        XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
                        hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

                        byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
                        if (logo != null && logo.length > 0) {
                                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
                        }

                        for (int row = 0; row <= 4; row++) {
                                UtilExcel.combinarCeldas(hoja, row, row, 1, 27);
                        }

                        CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
                        UtilExcel.establecerTexto(hoja, 0, 1,
                                        UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
                        UtilExcel.establecerTexto(hoja, 1, 1,
                                        "HORAS EXTRA - CONSOLIDADO", estiloTitulo);
                        UtilExcel.establecerTexto(hoja, 2, 1,
                                        "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL "
                                                        + safe(request.getFechaFin()),
                                        estiloTitulo);

                        Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
                        for (int c = 0; c < HEADERS.length; c++) {
                                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
                        }

                        CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
                        UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloHeader);
                        UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);

                        int filaDatosIni = FILA_ENCABEZADO + 1;
                        int filaAct = filaDatosIni;
                        int item = 1;

                        for (GrupoHorasExtraConsolidadoDT grupo : safeGrupos(request.getGrupos())) {
                                for (EmpleadoHorasExtraConsolidadoDTO emp : safeEmpleados(grupo.getEmpleados())) {
                                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();

                                        for (RegistroHorasExtraConsolidadoDTO reg : safeRegistros(emp.getTLaborado())) {
                                                List<DetalleHoraExtraDTO> detalles = safeDetalles(
                                                                reg.getDetalleHorasExtra());
                                                if (detalles.isEmpty()) {
                                                        detalles = new ArrayList<>();
                                                        detalles.add(new DetalleHoraExtraDTO());
                                                }

                                                for (DetalleHoraExtraDTO detalle : detalles) {
                                                        Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                                                        int col = 0;

                                                        UtilExcel.establecerValor(r, col++, item++, null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(emp.getIdentificacion()), null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getCodigo()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, apenom, null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getCiudad()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getSucursal()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getRegimen()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getDepartamento()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++, safe(emp.getCargo()), null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(reg.getEntrada() != null ? reg.getEntrada()
                                                                                        .getFecha_hora_horario() : ""),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        extraerHora(reg.getEntrada() != null ? reg
                                                                                        .getEntrada()
                                                                                        .getFecha_hora_horario()
                                                                                        : null),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        formatearTimbre(
                                                                                        reg.getEntrada() != null ? reg
                                                                                                        .getEntrada()
                                                                                                        .getFecha_hora_horario()
                                                                                                        : null,
                                                                                        reg.getEntrada() != null ? reg
                                                                                                        .getEntrada()
                                                                                                        .getFecha_hora_timbre()
                                                                                                        : null),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        extraerHora(reg.getInicioAlimentacion() != null
                                                                                        ? reg.getInicioAlimentacion()
                                                                                                        .getFecha_hora_horario()
                                                                                        : null),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        formatearTimbre(
                                                                                        reg.getInicioAlimentacion() != null
                                                                                                        ? reg.getInicioAlimentacion()
                                                                                                                        .getFecha_hora_horario()
                                                                                                        : null,
                                                                                        reg.getInicioAlimentacion() != null
                                                                                                        ? reg.getInicioAlimentacion()
                                                                                                                        .getFecha_hora_timbre()
                                                                                                        : null),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        extraerHora(reg.getFinAlimentacion() != null
                                                                                        ? reg.getFinAlimentacion()
                                                                                                        .getFecha_hora_horario()
                                                                                        : null),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        formatearTimbre(
                                                                                        reg.getFinAlimentacion() != null
                                                                                                        ? reg.getFinAlimentacion()
                                                                                                                        .getFecha_hora_horario()
                                                                                                        : null,
                                                                                        reg.getFinAlimentacion() != null
                                                                                                        ? reg.getFinAlimentacion()
                                                                                                                        .getFecha_hora_timbre()
                                                                                                        : null),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        extraerHora(reg.getSalida() != null ? reg
                                                                                        .getSalida()
                                                                                        .getFecha_hora_horario()
                                                                                        : null),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        formatearTimbre(
                                                                                        reg.getSalida() != null ? reg
                                                                                                        .getSalida()
                                                                                                        .getFecha_hora_horario()
                                                                                                        : null,
                                                                                        reg.getSalida() != null ? reg
                                                                                                        .getSalida()
                                                                                                        .getFecha_hora_timbre()
                                                                                                        : null),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(reg.getMinAtrasos()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(
                                                                                        reg.getMinSalidasAnticipadas()),
                                                                        null);

                                                        Double minAsignado = 0d;
                                                        if (reg.getInicioAlimentacion() != null && reg
                                                                        .getInicioAlimentacion()
                                                                        .getMinutos_alimentacion() != null) {
                                                                minAsignado = reg.getInicioAlimentacion()
                                                                                .getMinutos_alimentacion();
                                                        }

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(minAsignado), null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(
                                                                                        reg.getMinAlimentacion()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(reg.getMinLaborados()),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getHorasExtra()), null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        formatDouble(detalle.getMinutosHorasExtra()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getTipoPorcentaje()), null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getPorcentaje()), null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getTipoRecargo()), null);
                                                }
                                        }
                                }
                        }

                        int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

                        CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
                        CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

                        UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                                        0, HEADERS.length - 1, estiloCentroBorde, true);

                        if (ultimaFila >= filaDatosIni) {
                                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila,
                                                0, 0, estiloCentroBorde, true);
                                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila,
                                                1, HEADERS.length - 1, estiloIzqBorde, true);
                        }

                        boolean[] filtros = new boolean[HEADERS.length];
                        for (int i = 0; i < filtros.length; i++)
                                filtros[i] = true;
                        filtros[0] = false;

                        if (ultimaFila >= filaDatosIni) {
                                UtilExcel.crearTablaEstilizada(
                                                hoja,
                                                "HorasExtraConsolidadoTabla",
                                                FILA_ENCABEZADO, 0,
                                                ultimaFila, HEADERS.length - 1,
                                                true,
                                                filtros);
                        }

                        libro.write(baos);
                        return baos.toByteArray();

                } catch (IllegalArgumentException e) {
                        throw e;
                } catch (Exception e) {
                        throw new ReportBuildException("No se pudo generar HorasExtraConsolidado.xlsx", e);
                }
        }

        private List<GrupoHorasExtraConsolidadoDT> safeGrupos(List<GrupoHorasExtraConsolidadoDT> grupos) {
                return grupos == null ? new ArrayList<>() : grupos;
        }

        private List<EmpleadoHorasExtraConsolidadoDTO> safeEmpleados(List<EmpleadoHorasExtraConsolidadoDTO> empleados) {
                return empleados == null ? new ArrayList<>() : empleados;
        }

        private List<RegistroHorasExtraConsolidadoDTO> safeRegistros(List<RegistroHorasExtraConsolidadoDTO> registros) {
                return registros == null ? new ArrayList<>() : registros;
        }

        private List<DetalleHoraExtraDTO> safeDetalles(List<DetalleHoraExtraDTO> detalles) {
                return detalles == null ? new ArrayList<>() : detalles;
        }

        private String safe(Object v) {
                if (v == null)
                        return "";
                String s = String.valueOf(v).trim();
                return "null".equalsIgnoreCase(s) ? "" : s;
        }

        private double nz(Double v) {
                return v == null ? 0d : v;
        }

        private String formatDouble(Double value) {
                if (value == null)
                        return "0";
                if (Math.floor(value) == value) {
                        return String.valueOf(value.intValue());
                }
                return String.valueOf(value);
        }

        private String convertirMinutosATiempo(Double minutos) {
                if (minutos == null || minutos <= 0)
                        return "00:00:00";
                int totalSegundos = (int) Math.round(minutos * 60);
                int horas = totalSegundos / 3600;
                int mins = (totalSegundos % 3600) / 60;
                int segundos = totalSegundos % 60;
                return String.format("%02d:%02d:%02d", horas, mins, segundos);
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
}