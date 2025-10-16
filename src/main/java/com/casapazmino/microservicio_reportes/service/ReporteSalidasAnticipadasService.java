package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.ReporteSalidasAnticipadasRequest;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.GrupoSalidasDTO;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.EmpleadoSalidaDTO;
import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.SalidaDTO;
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
public class ReporteSalidasAnticipadasService {

    public byte[] generarReportePDF(ReporteSalidasAnticipadasRequest request) {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4, 40, 40, 30, 50);
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            String titulo = "SALIDAS ANTICIPADAS - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));
            document.add(ReporteUtil
                    .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados().forEach(emp -> contadorGlobal.addAndGet(emp.getSalidas().size())));

            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(new float[] { 8, 2 });
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            int contador = 1;
            for (GrupoSalidasDTO grupo : request.getGrupos()) {
                for (EmpleadoSalidaDTO emp : grupo.getEmpleados()) {

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

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    document.add(tablaContenedora);

                    PdfPTable tablaSalidas = new PdfPTable(12);
                    tablaSalidas.setWidthPercentage(100);
                    tablaSalidas.setSpacingBefore(5f);
                    tablaSalidas.setWidths(new float[] { 0.5f, 1.8f, 1.5f, 1.8f, 2, 2, 2, 2, 2, 2, 2, 2 });

                    PdfPCell h1 = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h1.setRowspan(2);
                    tablaSalidas.addCell(h1);
                    PdfPCell h2 = ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h2.setColspan(2);
                    tablaSalidas.addCell(h2);
                    PdfPCell h3 = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h3.setColspan(2);
                    tablaSalidas.addCell(h3);
                    PdfPCell h4 = ReporteUtil.crearCelda("TIPO PERMISO", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal);
                    h4.setRowspan(2);
                    tablaSalidas.addCell(h4);
                    PdfPCell h5 = ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h5.setRowspan(2);
                    tablaSalidas.addCell(h5);
                    PdfPCell h6 = ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h6.setRowspan(2);
                    tablaSalidas.addCell(h6);
                    PdfPCell h7 = ReporteUtil.crearCelda("PERMISO", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    h7.setColspan(2);
                    tablaSalidas.addCell(h7);
                    PdfPCell h8 = ReporteUtil.crearCelda("SALIDA ANTICIPADA", ReporteUtil.fuenteEncabezado(),
                            colorPrincipal);
                    h8.setColspan(2);
                    tablaSalidas.addCell(h8);

                    // Segunda fila encabezados (8 columnas)
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas
                            .addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaSalidas.addCell(ReporteUtil.crearCelda("", ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    long totalSegundos = 0;
                    double totalMinutos = 0;

                    for (SalidaDTO s : emp.getSalidas()) {

                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;
                        String[] horaHorario = s.getFecha_hora_horario().split(" ");
                        String[] horaTimbre = s.getFecha_hora_timbre().split(" ");

                        long segundos = s.getDiferencia() != null ? Math.round(s.getDiferencia()) : 0;
                        long horas = segundos / 3600;
                        long minutos = (segundos % 3600) / 60;
                        long restoSeg = segundos % 60;
                        String tiempoFormateado = String.format("%02d:%02d:%02d", horas, minutos, restoSeg);

                        tablaSalidas.addCell(
                                ReporteUtil.crearCelda(String.valueOf(contador), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(horaHorario[0]),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(horaHorario.length > 1 ? horaHorario[1] : "",
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(horaTimbre[0]),
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(horaTimbre.length > 1 ? horaTimbre[1] : "",
                                ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas
                                .addCell(ReporteUtil.crearCelda(s.getTipo_permiso(), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(s.getDesde(), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(s.getHasta(), ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(" ", ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas.addCell(ReporteUtil.crearCelda(" ", ReporteUtil.fuenteTexto(), fondo));
                        tablaSalidas
                                .addCell(ReporteUtil.crearCelda(tiempoFormateado, ReporteUtil.fuenteTexto(), fondo)); // hh:mm:ss
                        tablaSalidas.addCell(ReporteUtil.crearCelda(String.format("%.2f", s.getDiferencia() / 60.0),
                                ReporteUtil.fuenteTexto(), fondo));
                        contador++;
                        totalSegundos += segundos;
                        totalMinutos += s.getDiferencia() != null ? s.getDiferencia() / 60.0 : 0;

                    }

                    // 9 celdas vacías
                    for (int i = 0; i < 9; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER);
                        tablaSalidas.addCell(celdaVacia);
                    }

                    // Celda "TOTAL"
                    tablaSalidas.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE));

                    // Tiempo total formateado
                    long horasT = totalSegundos / 3600;
                    long minutosT = (totalSegundos % 3600) / 60;
                    long segRest = totalSegundos % 60;
                    String tiempoTotal = String.format("%02d:%02d:%02d", horasT, minutosT, segRest);
                    tablaSalidas.addCell(ReporteUtil.crearCelda(tiempoTotal, ReporteUtil.fuenteTexto(), Color.WHITE));

                    // Total en minutos
                    tablaSalidas.addCell(ReporteUtil.crearCelda(String.format("%.2f", totalMinutos),
                            ReporteUtil.fuenteTexto(), Color.WHITE));

                    document.add(tablaSalidas);
                    document.add(Chunk.NEWLINE);
                }
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
// XLSX (nuevo)
// =========================
public byte[] generarReporteXLSX(ReporteSalidasAnticipadasRequest request) {
    try (XSSFWorkbook libro = new XSSFWorkbook();
         ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

        XSSFSheet hoja = libro.createSheet("Salidas_Anticipadas");

        // 1) Logo estándar A1:B5
        byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
        UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

        // 2) Merges B1:O5  (B=1 .. O=14 en 0-based)
        for (int row = 0; row <= 4; row++) {
            UtilExcel.combinarCeldas(hoja, row, row, 1, 14);
        }

        // 3) Títulos
        CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
        UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
        UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE SALIDAS ANTICIPADAS", estiloTitulo);
        String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL " + safe(request.getFechaFin());
        UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

        // 4) Encabezados (fila 6 -> idx 5) + anchos
        final int filaEnc = 5;
        String[] headers = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "FECHA HORARIO", "HORA HORARIO",
                "FECHA TIMBRE", "HORA TIMBRE",
                "SALIDA ANTICIPADA HH:MM:SS", "SALIDA ANTICIPADA MINUTOS"
        };
        int[] anchos = { 10,20,20,20, 20,20,20,20,20, 20,20, 20,20, 20,20 };

        Row filaHeader = UtilExcel.asegurarFila(hoja, filaEnc);
        for (int c = 0; c < headers.length; c++) {
            UtilExcel.establecerTexto(filaHeader, c, headers[c], null);
        }
        CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
        UtilExcel.aplicarEstiloAFila(filaHeader, headers.length, estiloHeader);
        UtilExcel.establecerAnchosColumnas(hoja, anchos);
        hoja.getRow(filaEnc).setHeightInPoints(18f);

        // 5) Cuerpo (aplanado grupos → empleados → salidas)
        int filaDatosIni = filaEnc + 1;
        int filaAct = filaDatosIni;
        int item = 1;

        if (request.getGrupos() != null) {
            for (GrupoSalidasDTO grupo : request.getGrupos()) {
                if (grupo.getEmpleados() == null) continue;

                for (EmpleadoSalidaDTO usu : grupo.getEmpleados()) {
                    String apenom = (safe(usu.getApellido()) + " " + safe(usu.getNombre())).trim();

                    if (usu.getSalidas() == null) continue;

                    for (SalidaDTO sal : usu.getSalidas()) {
                        Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                        int col = 0;

                        // === Cálculos (como en TS) ===
                        String[] ph = splitFechaHora(sal.getFecha_hora_horario());
                        String[] pt = splitFechaHora(sal.getFecha_hora_timbre());

                        String horaHorario = ph[1];
                        String horaTimbre  = pt[1];

                        double minutos = segundosAMinutosConDecimales(sal.getDiferencia());
                        String tiempo   = convertirMinutosATiempo(minutos);

                        // === Escritura ===
                        UtilExcel.establecerValor(r, col++, item++, null);                // ITEM
                        UtilExcel.establecerTexto(r, col++, safe(usu.getIdentificacion()), null);
                        UtilExcel.establecerTexto(r, col++, safe(usu.getCodigo()), null);
                        UtilExcel.establecerTexto(r, col++, apenom, null);
                        UtilExcel.establecerTexto(r, col++, safe(usu.getCiudad()), null);
                        UtilExcel.establecerTexto(r, col++, safe(usu.getSucursal()), null);
                        UtilExcel.establecerTexto(r, col++, safe(usu.getRegimen()), null);
                        UtilExcel.establecerTexto(r, col++, safe(usu.getDepartamento()), null);
                        UtilExcel.establecerTexto(r, col++, safe(usu.getCargo()), null);

                        UtilExcel.establecerTexto(r, col++, ph[0], null);                 // FECHA HORARIO
                        UtilExcel.establecerTexto(r, col++, horaHorario, null);           // HORA HORARIO
                        UtilExcel.establecerTexto(r, col++, pt[0], null);                 // FECHA TIMBRE
                        UtilExcel.establecerTexto(r, col++, horaTimbre, null);            // HORA TIMBRE
                        UtilExcel.establecerTexto(r, col++, tiempo, null);                // HH:MM:SS
                        UtilExcel.establecerTexto(r, col++, String.format("%.2f", minutos), null); // minutos con decimales
                    }
                }
            }
        }

        int ultimaFila = (filaAct == filaDatosIni) ? filaEnc : (filaAct - 1);

        // 6) Estilos de cuerpo
        CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
        CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

        // Header centrado
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
            filtros[0] = false;

            UtilExcel.crearTablaEstilizada(
                    hoja,
                    "SalidaAnticipadaReporteTabla",
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

// ===== Helpers locales =====
private String safe(Object v) {
    if (v == null) return "";
    String s = String.valueOf(v).trim();
    return "null".equalsIgnoreCase(s) ? "" : s;
}

private String[] splitFechaHora(String fechaHora) {
    // Devuelve [fecha, hora] siempre
    if (fechaHora == null || !fechaHora.contains(" ")) return new String[] { "", "" };
    String[] p = fechaHora.split(" ");
    String fecha = p.length > 0 ? p[0] : "";
    String hora  = p.length > 1 ? p[1] : "";
    return new String[] { fecha, hora };
}

private double segundosAMinutosConDecimales(Double segundos) {
    if (segundos == null) return 0d;
    return segundos / 60.0;
}

private String convertirMinutosATiempo(Double minutos) {
    if (minutos == null || minutos <= 0) return "00:00:00";
    int totalSeg = (int) Math.round(minutos * 60);
    int h = totalSeg / 3600;
    int m = (totalSeg % 3600) / 60;
    int s = totalSeg % 60;
    return String.format("%02d:%02d:%02d", h, m, s);
}


}
