package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
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
                try {
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        Document document = new Document(PageSize.A4.rotate(), 40, 40, 30, 50);
                        PdfWriter writer = PdfWriter.getInstance(document, baos);
                        writer.setPageEvent(new ConfiguracionPaginaPDF(
                                        request.getUsuario(),
                                        request.getFraseMarcaAgua(),
                                        request.getColorPrincipal()));

                        document.open();

                        // Logo y encabezado
                        Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                        if (logo != null)
                                document.add(logo);

                        document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
                        String titulo = "REPORTE DE TIEMPO LABORADO - "
                                        + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
                        document.add(ReporteUtil.crearTituloReporte(titulo));
                        document.add(ReporteUtil
                                        .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL "
                                                        + request.getFechaFin()));

                        Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
                        Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
                        Color zebraColor = ReporteUtil.colorZebraClaro();

                        // Contador
                        AtomicInteger totalRegistros = new AtomicInteger();
                        request.getGrupos().forEach(
                                        grupo -> grupo.getEmpleados().forEach(emp -> {
                                                totalRegistros.addAndGet(emp.getTLaborado().size());
                                        }));

                        // Tabla de codigos de color
                        PdfPTable colores = new PdfPTable(5);
                        colores.setWidthPercentage(100);
                        colores.setWidths(new float[] { 3, 1.5f, 1.5f, 2, 2 });

                        colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0xEE4444)));
                        colores.addCell(ReporteUtil.celdaEncabezado("TIEMPO LABORADO MENOR AL PLANIFICADO",
                                        Color.WHITE));
                        colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0x55EE44)));

                        colores.setSpacingAfter(10f);
                        document.add(colores);

                        // Tabla de título y contador
                        PdfPTable tablaTitulo = new PdfPTable(2);
                        tablaTitulo.setWidthPercentage(100);
                        tablaTitulo.setWidths(new float[] { 8, 2 });
                        tablaTitulo.setSpacingAfter(10f);

                        PdfPCell celdaTitulo = new PdfPCell(
                                        new Phrase("LISTA DE EMPLEADOS", ReporteUtil.fuenteEncabezado()));
                        celdaTitulo.setBackgroundColor(colorSecundario);
                        celdaTitulo.setPadding(5f);
                        celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
                        tablaTitulo.addCell(celdaTitulo);

                        PdfPCell celdaContador = new PdfPCell(
                                        new Phrase("N° Registros: " + totalRegistros.get(),
                                                        ReporteUtil.fuenteEncabezado()));
                        celdaContador.setBackgroundColor(colorSecundario);
                        celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
                        celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        celdaContador.setPadding(5f);
                        celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
                        tablaTitulo.addCell(celdaContador);

                        document.add(tablaTitulo);

                        for (GrupoTiempoDTO grupo : request.getGrupos()) {
                                for (EmpleadoTiempoDTO emp : grupo.getEmpleados()) {

                                        // TABLA INFORMACION DEL EMPLEADO
                                        PdfPTable infoEmpleado = new PdfPTable(3);
                                        infoEmpleado.setWidthPercentage(100);
                                        infoEmpleado.setWidths(new float[] { 4, 4, 4 });

                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",
                                                        emp.getIdentificacion(), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                                        emp.getApellido() + " " + emp.getNombre(), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(),
                                                        zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:",
                                                        emp.getRegimen(), zebraColor));
                                        infoEmpleado
                                                        .addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:",
                                                                        emp.getDepartamento(), zebraColor));
                                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(),
                                                        zebraColor));

                                        PdfPTable tablaContenedora = new PdfPTable(1);
                                        tablaContenedora.setWidthPercentage(100);
                                        PdfPCell contenedor = new PdfPCell(infoEmpleado);
                                        contenedor.setPadding(0);
                                        contenedor.setBorder(Rectangle.BOX);
                                        tablaContenedora.addCell(contenedor);
                                        tablaContenedora.setSpacingAfter(5f);
                                        document.add(tablaContenedora);

                                        // TABLA DE ASISTENCIA POR DÍA (encabezado en una sola tabla con 16 columnas)
                                        PdfPTable encabezado = new PdfPTable(14);
                                        encabezado.setWidthPercentage(100);
                                        encabezado.setWidths(new float[] { 0.5f, 1.8f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.5f, 1.5f, 1.5f, 1.5f });
                                        // Fila 1 - encabezados agrupados con rowspan o colspan
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(),
                                                                        colorPrincipal, 2, 1));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(),
                                                                        colorPrincipal, 2, 1));

                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("ENTRADA",
                                                                        ReporteUtil.fuenteEncabezado(), colorPrincipal,
                                                                        1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("INICIO ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 1, 2));
                                        encabezado.addCell(ReporteUtil.crearCelda("FIN ALIMENTACIÓN",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 1, 2));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("SALIDA", ReporteUtil.fuenteEncabezado(),
                                                                        colorPrincipal, 1, 2));

                                        encabezado.addCell(ReporteUtil.crearCelda("TIEMPO PLANIFICADO",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 2, 2));

                                        encabezado.addCell(ReporteUtil.crearCelda("TIEMPO LABORADO",
                                                        ReporteUtil.fuenteEncabezado(),
                                                        colorPrincipal, 2, 2));

                                        // Fila 2 - subcolumnas debajo de agrupados
                                        for (int i = 0; i < 4; i++) {
                                                encabezado.addCell(
                                                                ReporteUtil.crearCelda("HORARIO",
                                                                                ReporteUtil.fuenteEncabezado(),
                                                                                colorPrincipal));
                                                encabezado.addCell(
                                                                ReporteUtil.crearCelda("TIMBRE",
                                                                                ReporteUtil.fuenteEncabezado(),
                                                                                colorSecundario));
                                        }

                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("MINUTOS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("HH:MM:SS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));

                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("MINUTOS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));
                                        encabezado.addCell(
                                                        ReporteUtil.crearCelda("HH:MM:SS",
                                                                        ReporteUtil.fuenteEncabezado(),
                                                                        colorSecundario));

                                        encabezado.setSpacingAfter(0f);
                                        document.add(encabezado);

                                        // TABLA DE DATOS
                                        PdfPTable tablaData = new PdfPTable(14);
                                        tablaData.setWidthPercentage(100);
                                        tablaData.setWidths(new float[] { 0.5f, 1.8f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.2f, 1.2f, 1.2f, 1.2f,
                                                        1.5f, 1.5f, 1.5f, 1.5f });

                                        Color colorFT = new Color(0xEE4444);
                                        Color colorTiempoMenorPlani = new Color(0x55EE44);

                                        int contador = 1;
                                        double totalPlanificadosMin = 0;
                                        double totalLaboradosMin = 0;

                                        for (RegistroTiempoDTO reg : emp.getTLaborado()) {
                                                Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                                                tablaData.addCell(
                                                                ReporteUtil.crearCelda(String.valueOf(contador),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                ReporteUtil.formatearFechaConDia(
                                                                                reg.getEntrada().getFecha_horario()),
                                                                ReporteUtil.fuenteTexto(), fondo));

                                                // ENTRADA
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                extraerHora(reg.getEntrada().getFecha_hora_horario()),
                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getEntrada()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getEntrada().getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg.getEntrada()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getEntrada().getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // INICIO ALIMENTACIÓN
                                                tablaData.addCell(
                                                                ReporteUtil.crearCelda(extraerHora(reg
                                                                                .getInicioAlimentacion()
                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getInicioAlimentacion()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getInicioAlimentacion()
                                                                                                .getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg
                                                                                                .getInicioAlimentacion()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getInicioAlimentacion()
                                                                                                                .getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // FIN ALIMENTACIÓN
                                                tablaData.addCell(
                                                                ReporteUtil.crearCelda(extraerHora(reg
                                                                                .getFinAlimentacion()
                                                                                .getFecha_hora_horario()),
                                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getFinAlimentacion()
                                                                                .getFecha_hora_horario(),
                                                                                reg.getFinAlimentacion()
                                                                                                .getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg.getFinAlimentacion()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getFinAlimentacion()
                                                                                                                .getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // SALIDA
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                extraerHora(reg.getSalida().getFecha_hora_horario()),
                                                                ReporteUtil.fuenteTexto(), fondo));
                                                tablaData.addCell(ReporteUtil.crearCelda(
                                                                formatearTimbre(reg.getSalida().getFecha_hora_horario(),
                                                                                reg.getSalida().getFecha_hora_timbre()),
                                                                ReporteUtil.fuenteTexto(),
                                                                getColorTimbre(
                                                                                formatearTimbre(reg.getSalida()
                                                                                                .getFecha_hora_horario(),
                                                                                                reg.getSalida().getFecha_hora_timbre()),
                                                                                fondo, colorFT)));

                                                // TIEMPO PLANIFICADO
                                                double minPlanificado = Double.parseDouble(reg.getMinPlanificados());
                                                double minLaborado = Double.parseDouble(reg.getMinLaborados());
                                                Color fondoTiempoLaborado = minLaborado < minPlanificado ? colorTiempoMenorPlani : fondo;

                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getTiempoPlanificado(), fondo));
                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getMinPlanificados(), fondo));

                                                // TIEMPO LABORADO
                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getTiempoLaborado(), fondoTiempoLaborado));
                                                tablaData.addCell(ReporteUtil.celdaCentro(reg.getMinLaborados(), fondoTiempoLaborado));

                                                
                                                contador++;
                                                
                                                totalPlanificadosMin += Double.parseDouble(reg.getMinPlanificados());
                                                totalLaboradosMin += Double.parseDouble(reg.getMinLaborados());

                                        }

                                        // Vacías hasta columna 10
                                        for (int i = 0; i < 9; i++) {
                                                PdfPCell celdaVacia = ReporteUtil.crearCelda("",
                                                                ReporteUtil.fuenteTexto(), Color.WHITE);
                                                celdaVacia.setBorder(Rectangle.NO_BORDER);
                                                tablaData.addCell(celdaVacia);
                                        }

                                        // Celda: TOTAL (Texto)
                                        tablaData.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(),
                                                        Color.WHITE));

                                        // Totales de TIEMPO PLANIFICADO
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        convertirMinutosATiempo(totalPlanificadosMin),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        String.format("%.2f", totalPlanificadosMin).replace(",", "."),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));

                                        // Totales de TIEMPO LABORADO
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        convertirMinutosATiempo(totalLaboradosMin),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));
                                        tablaData.addCell(ReporteUtil.crearCelda(
                                                        String.format("%.2f", totalLaboradosMin).replace(",", "."),
                                                        ReporteUtil.fuenteTexto(), Color.WHITE));

                                        tablaData.setSpacingAfter(10f);
                                        document.add(tablaData);

                                }

                        }

                        document.close();
                        return baos.toByteArray();

                } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                }
        }

        public byte[] generarReporteTiempoLaboradoExcel(ReporteTiempoLaboradoRequest request) {
                System.out.println("Generando XLSX de Tiempo Laborado...");
                try (XSSFWorkbook libro = new XSSFWorkbook();
                        ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

                        // =========================
                        // Hoja: Tiempo Laborado
                        // =========================
                        XSSFSheet hoja = libro.createSheet("Tiempo_laborado");

                        // 1) Logo estándar A1:B5
                        byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
                        UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

                        // 2) Merges B1:V5 (B=1 .. V=21 en 0-based)
                        for (int row = 0; row <= 4; row++) {
                        UtilExcel.combinarCeldas(hoja, row, row, 1, 21);
                        }

                        // 3) Títulos
                        CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
                        UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())), estiloTitulo);

                        String activosInactivos = ("1".equals(safe(request.getOpcionBusqueda())) ||
                                                "1".equals(String.valueOf(request.getOpcionBusqueda())))
                                                ? "ACTIVOS" : "INACTIVOS";
                        UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TIEMPO LABORADO - " + activosInactivos, estiloTitulo);

                        String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL " + safe(request.getFechaFin());
                        UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

                        // 4) Encabezados + anchos (fila 6 → idx 5)
                        final int filaEnc = 5;
                        String[] headers = {
                                "ITEM","IDENTIFICACIÓN","CÓDIGO","APELLIDO NOMBRE","CIUDAD","SUCURSAL",
                                "RÉGIMEN","DEPARTAMENTO","CARGO","FECHA",
                                "HORARIO ENTRADA","TIMBRE ENTRADA",
                                "HORARIO INICIO ALIMENTACIÓN","TIMBRE INICIO ALIMENTACIÓN",
                                "HORARIO FIN ALIMENTACIÓN","TIMBRE FIN ALIMENTACIÓN",
                                "HORARIO SALIDA","TIMBRE SALIDA",
                                "TIEMPO PLANIFICADO","TIEMPO PLANIFICADO MINUTOS",
                                "TIEMPO LABORADO","TIEMPO LABORADO MINUTOS"
                        };
                        int[] anchos = {
                                10,20,20,28, 18,18,18,20,18, 18,
                                18,18, 22,22, 22,22, 18,18, 22,26, 22,26
                        };

                        Row fh = UtilExcel.asegurarFila(hoja, filaEnc);
                        for (int c = 0; c < headers.length; c++) {
                        UtilExcel.establecerTexto(fh, c, headers[c], null);
                        }
                        CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
                        UtilExcel.aplicarEstiloAFila(fh, headers.length, estiloHeader);
                        UtilExcel.establecerAnchosColumnas(hoja, anchos);
                        hoja.getRow(filaEnc).setHeightInPoints(18f);

                        // 5) Cuerpo (aplanado grupos → empleados → tLaborado)
                        int filaDatosIni = filaEnc + 1;
                        int filaAct = filaDatosIni;
                        int item = 1;

                        if (request.getGrupos() != null) {
                        for (GrupoTiempoDTO grupo : request.getGrupos()) {
                                if (grupo == null || grupo.getEmpleados() == null) continue;

                                for (EmpleadoTiempoDTO emp : grupo.getEmpleados()) {
                                if (emp == null || emp.getTLaborado() == null) continue;

                                String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                                String ciudad   = firstNonEmpty(safe(emp.getCiudad()), safe(grupo.getCiudad()));
                                String sucursal = firstNonEmpty(safe(emp.getSucursal()), safe(grupo.getSucursal()));

                                for (RegistroTiempoDTO reg : emp.getTLaborado()) {
                                        if (reg == null) continue;

                                        boolean esEAS   = "EAS".equalsIgnoreCase(safe(reg.getTipo()));
                                        boolean control = "true".equalsIgnoreCase(String.valueOf(reg.getControl())) ||
                                                        Boolean.TRUE.equals(reg.getControl());

                                        // FECHA (tomamos la fecha_hora_horario de ENTRADA)
                                        String fecha = safe(() -> reg.getEntrada().getFecha_hora_horario());

                                        // HORARIOS (HH:mm:ss)
                                        String entradaHorario = horaDe(safe(() -> reg.getEntrada().getFecha_hora_horario()));
                                        String salidaHorario  = horaDe(safe(() -> reg.getSalida().getFecha_hora_horario()));
                                        String iaHorario      = esEAS ? horaDe(safe(() -> reg.getInicioAlimentacion().getFecha_hora_horario())) : "";
                                        String faHorario      = esEAS ? horaDe(safe(() -> reg.getFinAlimentacion().getFecha_hora_horario()))    : "";

                                        // TIMBRES (regla del frontend antiguo: hora || L/FD || FT/SCA según control)
                                        String origen = safe(reg.getOrigen());
                                        String entradaTimbre = toTimbre(
                                                safe(() -> reg.getEntrada().getFecha_hora_horario()),
                                                safe(() -> reg.getEntrada().getFecha_hora_timbre()),
                                                origen, control);

                                        String salidaTimbre = toTimbre(
                                                safe(() -> reg.getSalida().getFecha_hora_horario()),
                                                safe(() -> reg.getSalida().getFecha_hora_timbre()),
                                                origen, control);

                                        String iaTimbre = esEAS ? toTimbre(
                                                safe(() -> reg.getInicioAlimentacion().getFecha_hora_horario()),
                                                safe(() -> reg.getInicioAlimentacion().getFecha_hora_timbre()),
                                                origen, control) : "";

                                        String faTimbre = esEAS ? toTimbre(
                                                safe(() -> reg.getFinAlimentacion().getFecha_hora_horario()),
                                                safe(() -> reg.getFinAlimentacion().getFecha_hora_timbre()),
                                                origen, control) : "";

                                        // Tiempos y minutos (si !control, replica planificado)
                                        String tiempoPlan = safe(reg.getTiempoPlanificado());
                                        String minPlan    = normalize2(safe(reg.getMinPlanificados()));

                                        String tiempoLab  = control ? safe(reg.getTiempoLaborado())  : tiempoPlan;
                                        String minLab     = control ? normalize2(safe(reg.getMinLaborados())) : minPlan;

                                        // === Escritura de fila ===
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

                        int ultimaFila = (filaAct == filaDatosIni) ? filaEnc : (filaAct - 1);

                        // 6) Estilos cuerpo
                        CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
                        CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

                        // Header centrado con bordes
                        UtilExcel.aplicarEstiloARegion(hoja, filaEnc, filaEnc, 0, headers.length - 1, estiloCentroBorde, true);

                        if (ultimaFila >= filaDatosIni) {
                        // ITEM centrado
                        UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                        // resto izquierda
                        UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, headers.length - 1, estiloIzqBorde, true);
                        }

                        // 7) Tabla estilizada + filtros (ITEM sin filtro)
                        if (ultimaFila >= filaDatosIni) {
                        boolean[] filtros = new boolean[headers.length];
                        for (int i = 0; i < filtros.length; i++) filtros[i] = true;
                        filtros[0] = false; // ITEM sin filtro

                        UtilExcel.crearTablaEstilizada(
                                hoja,
                                "TiempoLaboradoTabla",
                                filaEnc, 0,
                                ultimaFila, headers.length - 1,
                                true,
                                filtros
                        );
                        }

                        // 8) Finalizar
                        libro.write(baos);
                        return baos.toByteArray();

                } catch (Exception e) {
                        e.printStackTrace();
                        return null;
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
        if (v == null) return "";
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
        if (fechaHora == null || fechaHora.trim().isEmpty()) return "";
        int idx = fechaHora.indexOf(' ');
        if (idx < 0 || idx + 1 >= fechaHora.length()) return "";
        return fechaHora.substring(idx + 1);
        }

        // Regla timbre (clonando la del frontend antiguo):
        // - Si existe timbre -> HH:mm:ss del timbre
        // - Si NO hay timbre:
        //      - Si origen == 'L' o 'FD' -> devuelve ese código
        //      - Si control == true -> "FT"
        //      - Si control == false -> "SCA"
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
        if (v == null || v.isBlank()) return "0.00";
        String s = v.replace(",", ".");
        try {
                double d = Double.parseDouble(s);
                return String.format(java.util.Locale.US, "%.2f", d);
        } catch (Exception e) {
                return s;
          }
        }
}
