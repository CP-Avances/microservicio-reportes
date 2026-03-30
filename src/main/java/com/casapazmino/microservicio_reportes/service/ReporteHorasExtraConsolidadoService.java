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
                final float[] WIDTHS_27 = {
                                0.7f, 1.8f, // N°, FECHA

                                1.3f, 1.3f, 0.9f, // ENTRADA
                                1.3f, 1.3f, 0.9f, // INICIO ALIMENTACIÓN
                                1.3f, 1.3f, 0.9f, // FIN ALIMENTACIÓN
                                1.3f, 1.3f, 0.9f, // SALIDA

                                1.3f, // ATRASO
                                1.5f, // SALIDA ANTICIPADA

                                1.3f, 1.3f, 1.3f, // T. ALIMENTACIÓN: ASIGNADO, TOMADO, EXCESO
                                1.5f, // TIEMPO PLANIFICADO
                                1.5f, // TIEMPO LABORADO

                                1.4f, // HORAS EXTRA
                                1.2f, // MINUTOS
                                1.5f, // TIPO %
                                1.4f, // PORCENTAJE
                                1.8f, // TIPO RECARGO

                                2.0f // OBSERVACIONES
                };

                final float[] WIDTHS_28 = {
                                0.7f, 1.8f, // N°, FECHA

                                1.3f, 1.3f, 0.9f, // ENTRADA
                                1.3f, 1.3f, 0.9f, // INICIO ALIMENTACIÓN
                                1.3f, 1.3f, 0.9f, // FIN ALIMENTACIÓN
                                1.3f, 1.3f, 0.9f, // SALIDA

                                1.3f, // ATRASO
                                1.5f, // SALIDA ANTICIPADA

                                1.3f, 1.3f, 1.3f, // T. ALIMENTACIÓN: ASIGNADO, TOMADO, EXCESO
                                1.5f, // TIEMPO PLANIFICADO
                                1.5f, // TIEMPO LABORADO

                                1.4f, // HORAS EXTRA
                                1.2f, // MINUTOS
                                1.5f, // TIPO %
                                1.4f, // PORCENTAJE
                                1.8f, // TIPO RECARGO

                                1.5f, // VALOR HE
                                2.0f // OBSERVACIONES
                };

                final Color COLOR_FALTA_TIMBRE = new Color(0xEE4444);
                final Color COLOR_ATRASO = new Color(0xEEE344);
                final Color COLOR_SALIDA_ANTICIPADA = new Color(0x4499EE);
                final Color COLOR_EXCESO_ALIMENTACION = new Color(0x55EE44);
                final Color COLOR_PERMISO = new Color(0xF6B26B);
                final Color COLOR_VACACIONES = new Color(0xD9B8FF);
                final Color COLOR_JUSTIFICACION_HORAS_EXTRAS = new Color(0x4DB6AC);

                boolean mostrarMonetizacion = request.isMostrarMonetizacion();

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

                        PdfPTable tablaLeyenda = new PdfPTable(7);
                        tablaLeyenda.setWidthPercentage(100);
                        tablaLeyenda.setSpacingBefore(5f);
                        tablaLeyenda.setSpacingAfter(8f);
                        tablaLeyenda.setWidths(new float[] { 1.6f, 1.5f, 1.5f, 2.0f, 2.2f, 1.4f, 1.6f });

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "CÓDIGO DE COLOR",
                                        ReporteUtil.fuenteEncabezadoCompacto(),
                                        colorSecundario));

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "FALTA TIMBRE",
                                        ReporteUtil.fuenteTextoCompacto(),
                                        COLOR_FALTA_TIMBRE));

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "ATRASO",
                                        ReporteUtil.fuenteTextoCompacto(),
                                        COLOR_ATRASO));

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "SALIDA ANTICIPADA",
                                        ReporteUtil.fuenteTextoCompacto(),
                                        COLOR_SALIDA_ANTICIPADA));

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "EXCESO DE ALIMENTACIÓN",
                                        ReporteUtil.fuenteTextoCompacto(),
                                        COLOR_EXCESO_ALIMENTACION));

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "PERMISO",
                                        ReporteUtil.fuenteTextoCompacto(),
                                        COLOR_PERMISO));

                        tablaLeyenda.addCell(ReporteUtil.crearCeldaCompacta(
                                        "VACACIONES",
                                        ReporteUtil.fuenteTextoCompacto(),
                                        COLOR_VACACIONES));

                        document.add(tablaLeyenda);

                        int totalColumnasPdf = mostrarMonetizacion ? 28 : 27;
                        float[] widthsPdf = mostrarMonetizacion ? WIDTHS_28 : WIDTHS_27;

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
                                        double totalExcesoAlimentacion = 0;
                                        double totalPlanificado = 0;
                                        double totalLaborado = 0;
                                        double totalHorasExtra = 0;
                                        double totalMonetizado = 0;

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

                                        PdfPTable encabezado = new PdfPTable(totalColumnasPdf);
                                        encabezado.setWidthPercentage(100);
                                        encabezado.setWidths(widthsPdf);

                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("N°",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("FECHA",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));

                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("ENTRADA",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 1, 3));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("INICIO ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 1, 3));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("FIN ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 1, 3));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("SALIDA",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 1, 3));

                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("ATRASO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("SALIDA ANTICIPADA",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("T. ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 1, 3));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIEMPO PLANIFICADO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIEMPO LABORADO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));

                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("HORAS EXTRA",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("MINUTOS",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIPO %",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("PORCENTAJE",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIPO RECARGO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));

                                        if (mostrarMonetizacion) {
                                                encabezado.addCell(ReporteUtil.crearCeldaCompacta("VALOR HE $",
                                                                ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal,
                                                                2, 1));
                                        }

                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("OBSERVACIONES",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal, 2, 1));

                                        // Subencabezados ENTRADA
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("HORARIO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIMBRE",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorSecundario));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("EST",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));

                                        // Subencabezados INICIO ALIMENTACIÓN
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("HORARIO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIMBRE",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorSecundario));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("EST",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));

                                        // Subencabezados FIN ALIMENTACIÓN
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("HORARIO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIMBRE",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorSecundario));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("EST",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));

                                        // Subencabezados SALIDA
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("HORARIO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIMBRE",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorSecundario));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("EST",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));

                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("ASIGNADO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("TOMADO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));
                                        encabezado.addCell(ReporteUtil.crearCeldaCompacta("EXCESO",
                                                        ReporteUtil.fuenteEncabezadoCompacto(), colorPrincipal));

                                        encabezado.setSpacingAfter(0f);
                                        document.add(encabezado);

                                        PdfPTable tablaData = new PdfPTable(totalColumnasPdf);
                                        tablaData.setWidthPercentage(100);
                                        tablaData.setWidths(widthsPdf);

                                        int contador = 1;

                                        for (RegistroHorasExtraConsolidadoDTO reg : safeRegistros(emp.getTLaborado())) {
                                                List<DetalleHoraExtraDTO> detalles = safeDetalles(
                                                                reg.getDetalleHorasExtra());

                                                if (detalles.isEmpty()) {
                                                        detalles = new ArrayList<>();
                                                        detalles.add(new DetalleHoraExtraDTO());
                                                }

                                                for (DetalleHoraExtraDTO detalle : detalles) {
                                                        //
                                                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;
                                                        boolean esEAS = "EAS".equalsIgnoreCase(safe(reg.getTipo()));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        String.valueOf(contador),
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        ReporteUtil.formatearFechaConDia(
                                                                                        safe(reg.getEntrada() != null
                                                                                                        ? reg.getEntrada()
                                                                                                                        .getFecha_horario()
                                                                                                        : "")),
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        // =========================
                                                        // ENTRADA
                                                        // =========================
                                                        String entradaHorario = extraerHora(
                                                                        reg.getEntrada() != null ? reg.getEntrada()
                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String entradaTimbre = obtenerTextoTimbre(reg.getEntrada());
                                                        String entradaEstado = safe(detalle.getEstadoEntradaReporte())
                                                                        .isEmpty()
                                                                                        ? obtenerTextoEstado(reg
                                                                                                        .getEntrada(),
                                                                                                        reg.getOrigen(),
                                                                                                        reg.getControl())
                                                                                        : safe(detalle.getEstadoEntradaReporte());

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        entradaHorario,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        entradaTimbre,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        entradaEstado,
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        getColorEstado(entradaEstado, fondo,
                                                                                        COLOR_FALTA_TIMBRE,
                                                                                        COLOR_PERMISO,
                                                                                        COLOR_VACACIONES,
                                                                                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                                                        // =========================
                                                        // INICIO ALIMENTACIÓN
                                                        // =========================
                                                        String iaHorario = esEAS
                                                                        ? extraerHora(reg
                                                                                        .getInicioAlimentacion() != null
                                                                                                        ? reg.getInicioAlimentacion()
                                                                                                                        .getFecha_hora_horario()
                                                                                                        : null)
                                                                        : "";

                                                        String iaTimbre = esEAS
                                                                        ? obtenerTextoTimbre(
                                                                                        reg.getInicioAlimentacion())
                                                                        : "";

                                                        String iaEstado = esEAS
                                                                        ? obtenerTextoEstado(
                                                                                        reg.getInicioAlimentacion(),
                                                                                        reg.getOrigen(),
                                                                                        reg.getControl())
                                                                        : "";

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        iaHorario,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        iaTimbre,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        iaEstado,
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        getColorEstado(iaEstado, fondo,
                                                                                        COLOR_FALTA_TIMBRE,
                                                                                        COLOR_PERMISO,
                                                                                        COLOR_VACACIONES,
                                                                                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                                                        // =========================
                                                        // FIN ALIMENTACIÓN
                                                        // =========================
                                                        String faHorario = esEAS
                                                                        ? extraerHora(reg.getFinAlimentacion() != null
                                                                                        ? reg.getFinAlimentacion()
                                                                                                        .getFecha_hora_horario()
                                                                                        : null)
                                                                        : "";

                                                        String faTimbre = esEAS
                                                                        ? obtenerTextoTimbre(reg.getFinAlimentacion())
                                                                        : "";

                                                        System.out.println("DEBUG DETALLE PDF >>> " +
                                                                        "fecha="
                                                                        + (reg.getEntrada() != null ? reg.getEntrada()
                                                                                        .getFecha_horario() : "")
                                                                        +
                                                                        ", tipoRecargo=" + detalle.getTipoRecargo() +
                                                                        ", estadoEntradaReporte="
                                                                        + detalle.getEstadoEntradaReporte() +
                                                                        ", estadoFinAlimentacionReporte="
                                                                        + detalle.getEstadoFinAlimentacionReporte());

                                                        String faEstado = esEAS
                                                                        ? (safe(detalle.getEstadoFinAlimentacionReporte())
                                                                                        .isEmpty()
                                                                                                        ? obtenerTextoEstado(
                                                                                                                        reg.getFinAlimentacion(),
                                                                                                                        reg.getOrigen(),
                                                                                                                        reg.getControl())
                                                                                                        : safe(detalle.getEstadoFinAlimentacionReporte()))
                                                                        : "";

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        faHorario,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        faTimbre,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        faEstado,
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        getColorEstado(faEstado, fondo,
                                                                                        COLOR_FALTA_TIMBRE,
                                                                                        COLOR_PERMISO,
                                                                                        COLOR_VACACIONES,
                                                                                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                                                        // =========================
                                                        // SALIDA
                                                        // =========================
                                                        String salidaHorario = extraerHora(
                                                                        reg.getSalida() != null ? reg.getSalida()
                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String salidaTimbre = obtenerTextoTimbre(reg.getSalida());
                                                        String salidaEstado = obtenerTextoEstado(reg.getSalida(),
                                                                        reg.getOrigen(), reg.getControl());

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        salidaHorario,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        salidaTimbre,
                                                                        ReporteUtil.fuenteTextoCompacto(), fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        salidaEstado,
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        getColorEstado(salidaEstado, fondo,
                                                                                        COLOR_FALTA_TIMBRE,
                                                                                        COLOR_PERMISO,
                                                                                        COLOR_VACACIONES,
                                                                                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));
                                                        //

                                                        ///////
                                                        boolean atrasoEsJHE = "JHE".equalsIgnoreCase(entradaEstado);

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        convertirMinutosATiempo(reg.getMinAtrasos()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        reg.getMinAtrasos() != null
                                                                                        && reg.getMinAtrasos() > 0
                                                                                        && !atrasoEsJHE
                                                                                                        ? COLOR_ATRASO
                                                                                                        : fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        convertirMinutosATiempo(
                                                                                        reg.getMinSalidasAnticipadas()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        reg.getMinSalidasAnticipadas() != null && reg
                                                                                        .getMinSalidasAnticipadas() > 0
                                                                                                        ? COLOR_SALIDA_ANTICIPADA
                                                                                                        : fondo));

                                                        Double minAsignado = 0d;
                                                        if (esEAS && reg.getInicioAlimentacion() != null
                                                                        && reg.getInicioAlimentacion()
                                                                                        .getMinutos_alimentacion() != null) {
                                                                minAsignado = reg.getInicioAlimentacion()
                                                                                .getMinutos_alimentacion();
                                                        }

                                                        Double minAlimentacion = reg.getMinAlimentacion() != null
                                                                        ? reg.getMinAlimentacion()
                                                                        : 0d;

                                                        Double minExcesoAlimentacion = calcularExcesoAlimentacion(
                                                                        minAsignado, minAlimentacion);

                                                        Double minPlanificadosBase = reg.getMinPlanificados() != null
                                                                        ? reg.getMinPlanificados()
                                                                        : 0d;
                                                        Double minPlanificados = minPlanificadosBase
                                                                        - (minAsignado != null ? minAsignado : 0d);

                                                        if (minPlanificados < 0) {
                                                                minPlanificados = 0d;
                                                        }

                                                        boolean alimentacionEsPermiso = "P".equalsIgnoreCase(iaEstado)
                                                                        || "P".equalsIgnoreCase(faEstado);
                                                        boolean alimentacionEsJHE = "JHE".equalsIgnoreCase(faEstado);
                                                        ///////////
                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        esEAS ? convertirMinutosATiempo(minAsignado)
                                                                                        : "",
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        esEAS ? convertirMinutosATiempo(minAlimentacion)
                                                                                        : "",
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        (esEAS
                                                                                        && !alimentacionEsPermiso
                                                                                        && minAsignado != null
                                                                                        && !alimentacionEsJHE
                                                                                        && minAlimentacion != null
                                                                                        && minAlimentacion > minAsignado)
                                                                                                        ? COLOR_EXCESO_ALIMENTACION
                                                                                                        : fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        esEAS ? convertirMinutosATiempo(
                                                                                        minExcesoAlimentacion) : "",
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        (esEAS
                                                                                        && !alimentacionEsPermiso
                                                                                        && !alimentacionEsJHE
                                                                                        && minExcesoAlimentacion != null
                                                                                        && minExcesoAlimentacion > 0)
                                                                                                        ? COLOR_EXCESO_ALIMENTACION
                                                                                                        : fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        convertirMinutosATiempo(minPlanificados),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        convertirMinutosATiempo(reg.getMinLaborados()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));
                                                        /////////
                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        safe(detalle.getHorasExtra()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        formatDouble(detalle.getMinutosHorasExtra()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        safe(detalle.getTipoPorcentaje()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        safe(detalle.getPorcentaje()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                        safe(detalle.getTipoRecargo()),
                                                                        ReporteUtil.fuenteTextoCompacto(),
                                                                        fondo));

                                                        if (mostrarMonetizacion) {
                                                                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                                formatMoney(detalle.getTotalAPagar()),
                                                                                ReporteUtil.fuenteTextoCompacto(),
                                                                                fondo));
                                                        }

                                                        tablaData.addCell(ReporteUtil.crearCeldaObservacionConEstado(
                                                                        safe(reg.getObservaciones()),
                                                                        fondo));

                                                
                                                       totalAtrasos += atrasoEsJHE ? 0d : nz(reg.getMinAtrasos());
                                                        totalSalidasAnticipadas += nz(reg.getMinSalidasAnticipadas());
                                                        totalAlimentacionTomado += esEAS ? nz(minAlimentacion) : 0d;
                                                        totalAlimentacionAsignado += esEAS ? nz(minAsignado) : 0d;
                                                        totalExcesoAlimentacion += (esEAS && !alimentacionEsJHE)
                                                                        ? nz(minExcesoAlimentacion)
                                                                        : 0d;
                                                        totalPlanificado += nz(minPlanificados);
                                                        totalLaborado += nz(reg.getMinLaborados());
                                                        totalHorasExtra += nz(detalle.getMinutosHorasExtra());
                                                        totalMonetizado += nz(detalle.getTotalAPagar());
                                                        contador++;

                                                        //////
                                                }
                                        }

                                        ///////////////////////
                                        int columnasVaciasTotal = 13;

                                        // ===============================
                                        // FILA 1 Y 2 CON CELDA TOTAL rowspan=2
                                        // ===============================

                                        // FILA 1: vacías antes de TOTAL
                                        for (int i = 0; i < columnasVaciasTotal; i++) {
                                                PdfPCell vacia = ReporteUtil.crearCeldaCompacta(
                                                                "",
                                                                ReporteUtil.fuenteTextoCompacto(),
                                                                Color.WHITE);
                                                vacia.setBorder(Rectangle.NO_BORDER);
                                                tablaData.addCell(vacia);
                                        }

                                        // Celda TOTAL ocupando 2 filas
                                        PdfPCell celdaTotal = ReporteUtil.crearCeldaCompacta(
                                                        "TOTAL",
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE);
                                        celdaTotal.setRowspan(2);
                                        celdaTotal.setNoWrap(true);
                                        celdaTotal.setHorizontalAlignment(Element.ALIGN_CENTER);
                                        celdaTotal.setVerticalAlignment(Element.ALIGN_MIDDLE);
                                        tablaData.addCell(celdaTotal);

                                        // Completar resto de la FILA 1 con vacías
                                        int columnasDespuesDeTotal = mostrarMonetizacion ? 14 : 13;
                                        for (int i = 0; i < columnasDespuesDeTotal; i++) {
                                                PdfPCell vacia = ReporteUtil.crearCeldaCompacta(
                                                                "",
                                                                ReporteUtil.fuenteTextoCompacto(),
                                                                Color.WHITE);
                                                vacia.setBorder(Rectangle.NO_BORDER);
                                                tablaData.addCell(vacia);
                                        }

                                        // ===============================
                                        // FILA 2: vacías antes de ATRASO
                                        // ===============================
                                        for (int i = 0; i < columnasVaciasTotal; i++) {
                                                PdfPCell vacia = ReporteUtil.crearCeldaCompacta(
                                                                "",
                                                                ReporteUtil.fuenteTextoCompacto(),
                                                                Color.WHITE);
                                                vacia.setBorder(Rectangle.NO_BORDER);
                                                tablaData.addCell(vacia);
                                        }

                                        // ATRASO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalAtrasos),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // SALIDA ANTICIPADA
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalSalidasAnticipadas),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // ASIGNADO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalAlimentacionAsignado),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // TOMADO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalAlimentacionTomado),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // EXCESO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalExcesoAlimentacion),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // TIEMPO PLANIFICADO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalPlanificado),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // TIEMPO LABORADO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalLaborado),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // HORAS EXTRA
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        convertirMinutosATiempo(totalHorasExtra),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // MINUTOS
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        formatDouble(totalHorasExtra),
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // TIPO %
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        "",
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // PORCENTAJE
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        "",
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        // TIPO RECARGO
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        "",
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));

                                        if (mostrarMonetizacion) {
                                                // VALOR HE
                                                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                                formatMoney(totalMonetizado),
                                                                ReporteUtil.fuenteTextoCompacto(),
                                                                Color.WHITE));
                                        }

                                        // OBSERVACIONES
                                        tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                                                        "",
                                                        ReporteUtil.fuenteTextoCompacto(),
                                                        Color.WHITE));
                                        //// ///////
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

                final boolean mostrarMonetizacion = request.isMostrarMonetizacion();

                final String[] HEADERS_35 = {
                                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL", "RÉGIMEN",
                                "DEPARTAMENTO", "CARGO", "FECHA",

                                "HORARIO ENTRADA", "TIMBRE ENTRADA", "EST ENTRADA",
                                "HORARIO INICIO ALIMENTACIÓN", "TIMBRE INICIO ALIMENTACIÓN", "EST INICIO ALIMENTACIÓN",
                                "HORARIO FIN ALIMENTACIÓN", "TIMBRE FIN ALIMENTACIÓN", "EST FIN ALIMENTACIÓN",
                                "HORARIO SALIDA", "TIMBRE SALIDA", "EST SALIDA",

                                "ATRASO", "SALIDA ANTICIPADA",
                                "TIEMPO ALIMENTACIÓN ASIGNADO", "TIEMPO ALIMENTACIÓN",
                                "TIEMPO ALIMENTACIÓN EXCESO",
                                "TIEMPO PLANIFICADO",
                                "TIEMPO LABORADO",

                                "HORAS EXTRA", "MINUTOS", "TIPO %", "PORCENTAJE", "TIPO RECARGO",
                                "OBSERVACIONES"
                };

                final String[] HEADERS_36 = {
                                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL", "RÉGIMEN",
                                "DEPARTAMENTO", "CARGO", "FECHA",

                                "HORARIO ENTRADA", "TIMBRE ENTRADA", "EST ENTRADA",
                                "HORARIO INICIO ALIMENTACIÓN", "TIMBRE INICIO ALIMENTACIÓN", "EST INICIO ALIMENTACIÓN",
                                "HORARIO FIN ALIMENTACIÓN", "TIMBRE FIN ALIMENTACIÓN", "EST FIN ALIMENTACIÓN",
                                "HORARIO SALIDA", "TIMBRE SALIDA", "EST SALIDA",

                                "ATRASO", "SALIDA ANTICIPADA",
                                "TIEMPO ALIMENTACIÓN ASIGNADO", "TIEMPO ALIMENTACIÓN",
                                "TIEMPO ALIMENTACIÓN EXCESO",
                                "TIEMPO PLANIFICADO",
                                "TIEMPO LABORADO",

                                "HORAS EXTRA", "MINUTOS", "TIPO %", "PORCENTAJE", "TIPO RECARGO",
                                "VALOR HE",
                                "OBSERVACIONES"
                };

                final String[] HEADERS = mostrarMonetizacion ? HEADERS_36 : HEADERS_35;

                final int[] ANCHOS_35 = {
                                10, 18, 14, 24, 18, 18, 18,
                                20, 18, 18,

                                18, 18, 12,
                                22, 22, 12,
                                22, 22, 12,
                                18, 18, 12,

                                16, 18,
                                22, 22, 22,
                                18,
                                18,

                                18, 14, 20, 16, 22,
                                28
                };

                final int[] ANCHOS_36 = {
                                10, 18, 14, 24, 18, 18, 18,
                                20, 18, 18,

                                18, 18, 12,
                                22, 22, 12,
                                22, 22, 12,
                                18, 18, 12,

                                16, 18,
                                22, 22, 22,
                                18,
                                18,

                                18, 14, 20, 16, 22,
                                18,
                                28
                };

                final int[] ANCHOS = mostrarMonetizacion ? ANCHOS_36 : ANCHOS_35;

                try (XSSFWorkbook libro = new XSSFWorkbook();
                                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

                        XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
                        hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

                        byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
                        if (logo != null && logo.length > 0) {
                                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
                        }

                        for (int row = 0; row <= 4; row++) {
                                UtilExcel.combinarCeldas(hoja, row, row, 1, HEADERS.length - 1);
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

                                                        boolean esEAS = "EAS".equalsIgnoreCase(safe(reg.getTipo()));

                                                        String entradaHorario = extraerHora(
                                                                        reg.getEntrada() != null
                                                                                        ? reg.getEntrada()
                                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String entradaTimbre = obtenerTextoTimbre(reg.getEntrada());
                                                        String entradaEstado = safe(detalle.getEstadoEntradaReporte())
                                                                        .isEmpty()
                                                                                        ? obtenerTextoEstado(reg
                                                                                                        .getEntrada(),
                                                                                                        reg.getOrigen(),
                                                                                                        reg.getControl())
                                                                                        : safe(detalle.getEstadoEntradaReporte());

                                                        String iaHorario = esEAS
                                                                        ? extraerHora(reg
                                                                                        .getInicioAlimentacion() != null
                                                                                                        ? reg.getInicioAlimentacion()
                                                                                                                        .getFecha_hora_horario()
                                                                                                        : null)
                                                                        : "";
                                                        String iaTimbre = esEAS
                                                                        ? obtenerTextoTimbre(
                                                                                        reg.getInicioAlimentacion())
                                                                        : "";
                                                        String iaEstado = esEAS
                                                                        ? obtenerTextoEstado(
                                                                                        reg.getInicioAlimentacion(),
                                                                                        reg.getOrigen(),
                                                                                        reg.getControl())
                                                                        : "";

                                                        String faHorario = esEAS
                                                                        ? extraerHora(reg.getFinAlimentacion() != null
                                                                                        ? reg.getFinAlimentacion()
                                                                                                        .getFecha_hora_horario()
                                                                                        : null)
                                                                        : "";
                                                        String faTimbre = esEAS
                                                                        ? obtenerTextoTimbre(reg.getFinAlimentacion())
                                                                        : "";
                                                        String faEstado = esEAS
                                                                        ? (safe(detalle.getEstadoFinAlimentacionReporte())
                                                                                        .isEmpty()
                                                                                                        ? obtenerTextoEstado(
                                                                                                                        reg.getFinAlimentacion(),
                                                                                                                        reg.getOrigen(),
                                                                                                                        reg.getControl())
                                                                                                        : safe(detalle.getEstadoFinAlimentacionReporte()))
                                                                        : "";

                                                        String salidaHorario = extraerHora(
                                                                        reg.getSalida() != null
                                                                                        ? reg.getSalida()
                                                                                                        .getFecha_hora_horario()
                                                                                        : null);
                                                        String salidaTimbre = obtenerTextoTimbre(reg.getSalida());
                                                        String salidaEstado = obtenerTextoEstado(
                                                                        reg.getSalida(),
                                                                        reg.getOrigen(),
                                                                        reg.getControl());

                                                        Double minAsignado = 0d;
                                                        if (esEAS
                                                                        && reg.getInicioAlimentacion() != null
                                                                        && reg.getInicioAlimentacion()
                                                                                        .getMinutos_alimentacion() != null) {
                                                                minAsignado = reg.getInicioAlimentacion()
                                                                                .getMinutos_alimentacion();
                                                        }
                                                        Double minAlimentacion = reg.getMinAlimentacion() != null
                                                                        ? reg.getMinAlimentacion()
                                                                        : 0d;

                                                        Double minExcesoAlimentacion = calcularExcesoAlimentacion(
                                                                        minAsignado, minAlimentacion);

                                                        Double minPlanificadosBase = reg.getMinPlanificados() != null
                                                                        ? reg.getMinPlanificados()
                                                                        : 0d;
                                                        Double minPlanificados = minPlanificadosBase
                                                                        - (minAsignado != null ? minAsignado : 0d);

                                                        if (minPlanificados < 0) {
                                                                minPlanificados = 0d;
                                                        }

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
                                                                        safe(reg.getEntrada() != null
                                                                                        ? reg.getEntrada()
                                                                                                        .getFecha_horario()
                                                                                        : ""),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++, entradaHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, entradaTimbre, null);
                                                        UtilExcel.establecerTexto(r, col++, entradaEstado, null);

                                                        UtilExcel.establecerTexto(r, col++, iaHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, iaTimbre, null);
                                                        UtilExcel.establecerTexto(r, col++, iaEstado, null);

                                                        UtilExcel.establecerTexto(r, col++, faHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, faTimbre, null);
                                                        UtilExcel.establecerTexto(r, col++, faEstado, null);

                                                        UtilExcel.establecerTexto(r, col++, salidaHorario, null);
                                                        UtilExcel.establecerTexto(r, col++, salidaTimbre, null);
                                                        UtilExcel.establecerTexto(r, col++, salidaEstado, null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(reg.getMinAtrasos()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(
                                                                                        reg.getMinSalidasAnticipadas()),
                                                                        null);
                                                        //////
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        esEAS ? convertirMinutosATiempo(minAsignado)
                                                                                        : "",
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        esEAS ? convertirMinutosATiempo(minAlimentacion)
                                                                                        : "",
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        esEAS ? convertirMinutosATiempo(
                                                                                        minExcesoAlimentacion) : "",
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(minPlanificados),
                                                                        null);

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        convertirMinutosATiempo(reg.getMinLaborados()),
                                                                        null);
                                                        ///////////
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getHorasExtra()), null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        formatDouble(detalle.getMinutosHorasExtra()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getTipoPorcentaje()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getPorcentaje()),
                                                                        null);
                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(detalle.getTipoRecargo()),
                                                                        null);

                                                        if (mostrarMonetizacion) {
                                                                UtilExcel.establecerTexto(r, col++,
                                                                                formatMoney(detalle.getTotalAPagar()),
                                                                                null);
                                                        }

                                                        UtilExcel.establecerTexto(r, col++,
                                                                        safe(reg.getObservaciones()),
                                                                        null);
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
                        for (int i = 0; i < filtros.length; i++) {
                                filtros[i] = true;
                        }
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

        private String formatMoney(Double value) {
                if (value == null) {
                        return "0.0000";
                }
                return String.format(java.util.Locale.US, "%.4f", value);
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

        private String obtenerTextoEstado(MarcaHorasExtraDTO marca, String origen, Boolean control) {
                if (marca == null) {
                        return "";
                }

                String estadoTimbre = safe(marca.getEstado_timbre());

                if (!estadoTimbre.isEmpty()) {
                        return estadoTimbre;
                }

                if ("L".equals(origen) || "FD".equals(origen) || "DHA".equals(origen)) {
                        return origen;
                }

                return (control != null && control) ? "FT" : "SCA";
        }

        private Color getColorEstado(
                        String valor,
                        Color porDefecto,
                        Color colorFT,
                        Color colorPermiso,
                        Color colorVacaciones,
                        Color colorJHE) {

                String estado = safe(valor);

                if ("FT".equalsIgnoreCase(estado)) {
                        return colorFT;
                }
                if ("P".equalsIgnoreCase(estado)) {
                        return colorPermiso;
                }
                if ("V".equalsIgnoreCase(estado)) {
                        return colorVacaciones;
                }
                if ("JHE".equalsIgnoreCase(estado)) {
                        return colorJHE;
                }
                return porDefecto;
        }

        private String obtenerTextoTimbre(MarcaHorasExtraDTO marca) {
                if (marca == null) {
                        return "";
                }
                if (marca.getFecha_hora_timbre() != null && !marca.getFecha_hora_timbre().trim().isEmpty()) {
                        return extraerHora(marca.getFecha_hora_timbre());
                }
                return "";
        }

        private Double calcularExcesoAlimentacion(Double minutosAsignados, Double minutosTomados) {
                double asignado = minutosAsignados != null ? minutosAsignados : 0d;
                double tomado = minutosTomados != null ? minutosTomados : 0d;

                if (tomado > asignado) {
                        return tomado - asignado;
                }

                return 0d;
        }

}