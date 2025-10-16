package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion.*;
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
public class ReporteTiempoAlimentacionService {

    public byte[] generarReporteTiempoAlimentacionPDF(ReporteTiempoAlimentacionRequest request) {
        System.out.println("Generando PDF de Tiempo de Alimentación...");
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

            String titulo = "TIEMPO DE ALIMENTACIÓN - "
                    + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
            document.add(ReporteUtil.crearTituloReporte(titulo));

            document.add(ReporteUtil
                    .crearTituloPeriodo("PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin()));
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();
            Color colorExceso = new Color(0x55EE44);
            Color colorFT = new Color(0xEE4444);

            PdfPTable colores = new PdfPTable(5);
            colores.setWidthPercentage(100);
            colores.setWidths(new float[] { 3, 1.5f, 1.5f, 2, 2 });

            colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0xEE4444)));
            colores.addCell(ReporteUtil.celdaEncabezado("EXCESO DE ALIMENTACIÓN", Color.WHITE));
            colores.addCell(ReporteUtil.celdaEncabezado(" ", new Color(0x55EE44)));

            colores.setSpacingAfter(10f);
            document.add(colores);

            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getGrupos().forEach(
                    grupo -> grupo.getEmpleados()
                            .forEach(emp -> contadorGlobal.addAndGet(emp.getAlimentacion().size())));

            PdfPTable tablaTitulo = new PdfPTable(2);
            tablaTitulo.setWidthPercentage(100);
            tablaTitulo.setWidths(new float[] { 8, 2 });

            PdfPCell celda1 = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celda1.setBackgroundColor(colorSecundario);
            celda1.setPadding(5f);
            celda1.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tablaTitulo.addCell(celda1);

            PdfPCell celda2 = new PdfPCell(
                    new Phrase("Nº Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celda2.setBackgroundColor(colorSecundario);
            celda2.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celda2.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celda2.setPadding(5);
            celda2.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tablaTitulo.addCell(celda2);

            tablaTitulo.setSpacingAfter(10f);
            document.add(tablaTitulo);

            for (GrupoAlimentacionDTO grupo : request.getGrupos()) {
                for (EmpleadoAlimentacionDTO emp : grupo.getEmpleados()) {

                    double totalExcesoAlimentacion = 0;
                    int contador = 1;

                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(new float[] { 4, 4, 4 });

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

                    PdfPTable tablaAlimentacion = new PdfPTable(7);
                    tablaAlimentacion.setWidthPercentage(100);
                    tablaAlimentacion.setWidths(new float[] { 0.5f, 1.8f, 1.8f, 1.8f, 1.8f, 1.8f, 1.8f });

                    String[] headers = { "Nº", "FECHA", "INICIO ALIMENTACIÓN", "FIN ALIMENTACIÓN", "M. ALIMENTACIÓN",
                            "M. TOMADOS", "M. EXCESO" };
                    for (String h : headers) {
                        tablaAlimentacion
                                .addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    }

                    for (RegistroAlimentacionDTO registro : emp.getAlimentacion()) {
                        Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                        String inicio = (registro.getInicioAlimentacion() == null
                                || registro.getInicioAlimentacion().isEmpty()) ? "FT"
                                        : extraerHora(registro.getInicioAlimentacion());
                        String fin = (registro.getFinAlimentacion() == null || registro.getFinAlimentacion().isEmpty())
                                ? "FT"
                                : extraerHora(registro.getFinAlimentacion());

                        Color colorInicio = "FT".equals(inicio) ? colorFT : fondo;
                        Color colorFin = "FT".equals(fin) ? colorFT : fondo;

                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(String.valueOf(contador), fondo));
                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(registro.getFecha(), fondo));
                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(inicio, colorInicio));
                        tablaAlimentacion.addCell(ReporteUtil.celdaCentro(fin, colorFin));
                        tablaAlimentacion.addCell(
                                ReporteUtil.celdaCentro(String.valueOf(registro.getMinutosPermitidos()), fondo));
                        tablaAlimentacion.addCell(
                                ReporteUtil.celdaCentro(String.valueOf(registro.getMinutosTomados()).replace(",", "."),
                                        fondo));
                        Color fondoExceso = registro.getMinutosExceso() > 0 ? colorExceso : fondo;
                        tablaAlimentacion.addCell(
                                ReporteUtil.celdaCentro(
                                        String.format("%.2f", registro.getMinutosExceso()).replace(",", "."),
                                        fondoExceso));

                        totalExcesoAlimentacion += registro.getMinutosExceso();
                        contador++;
                    }

                    for (int i = 0; i < 5; i++) {
                        PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                        celdaVacia.setBorder(Rectangle.NO_BORDER);
                        tablaAlimentacion.addCell(celdaVacia);
                    }
                    tablaAlimentacion.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE));
                    tablaAlimentacion.addCell(
                            ReporteUtil.crearCelda(String.format("%.2f", totalExcesoAlimentacion).replace(",", "."),
                                    ReporteUtil.fuenteTexto(), Color.WHITE));

                    tablaAlimentacion.setSpacingAfter(10f);
                    document.add(tablaAlimentacion);
                }
            }

            document.close();
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
    

    private String extraerHora(String fechaHora) {
        if (fechaHora == null || !fechaHora.contains(" "))
            return fechaHora;
        return fechaHora.split(" ")[1];
    }

    public byte[] generarReporteTiempoAlimentacionXLSX(ReporteTiempoAlimentacionRequest request) {
        System.out.println("Generando XLSX de Tiempo de Alimentación...");
        try (XSSFWorkbook libro = new XSSFWorkbook();
         ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

        // =========================
        // Hoja: Detalle de alimentación
        // =========================
        XSSFSheet hoja = libro.createSheet("Tiempo_Alimentacion");

        // 1) Logo estándar A1:B5
        byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
        UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

        // 2) Merges B1:O5 (B=1 .. O=14 en 0-based)
        for (int row = 0; row <= 4; row++) {
            UtilExcel.combinarCeldas(hoja, row, row, 1, 14);
        }

        // 3) Títulos
        CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
        UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())), estiloTitulo);

        String activosInactivos = ("1".equals(safe(request.getOpcionBusqueda())) || "1".equals(String.valueOf(request.getOpcionBusqueda())))
                ? "ACTIVOS" : "INACTIVOS";
        UtilExcel.establecerTexto(hoja, 1, 1,
                "TIEMPO DE ALIMENTACIÓN - " + activosInactivos,
                estiloTitulo);

        String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL " + safe(request.getFechaFin());
        UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

        // 4) Encabezados + anchos (fila 6 → idx 5)
        final int filaEnc = 5;
        String[] headers = {
                "ITEM","IDENTIFICACIÓN","CÓDIGO","APELLIDO NOMBRE",
                "CIUDAD","SUCURSAL","RÉGIMEN","DEPARTAMENTO","CARGO",
                "FECHA","INICIO ALIMENTACIÓN","FIN ALIMENTACIÓN",
                "MIN. PERMITIDOS","MIN. TOMADOS","MIN. EXCESO"
        };
        int[] anchos = {
                10,20,20,28, 18,18,18,20,18, 16,22,22, 18,18,18
        };

        Row fh = UtilExcel.asegurarFila(hoja, filaEnc);
        for (int c = 0; c < headers.length; c++) {
            UtilExcel.establecerTexto(fh, c, headers[c], null);
        }
        CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
        UtilExcel.aplicarEstiloAFila(fh, headers.length, estiloHeader);
        UtilExcel.establecerAnchosColumnas(hoja, anchos);
        hoja.getRow(filaEnc).setHeightInPoints(18f);

        // 5) Cuerpo (aplanado grupos → empleados → alimentacion)
        int filaDatosIni = filaEnc + 1;
        int filaAct = filaDatosIni;
        int item = 1;

        if (request.getGrupos() != null) {
            for (GrupoAlimentacionDTO grupo : request.getGrupos()) {
                if (grupo == null || grupo.getEmpleados() == null) continue;

                for (EmpleadoAlimentacionDTO emp : grupo.getEmpleados()) {
                    if (emp == null) continue;

                    String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();

                    if (emp.getAlimentacion() == null) continue;
                    for (RegistroAlimentacionDTO reg : emp.getAlimentacion()) {
                        if (reg == null) continue;

                        Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                        int col = 0;

                        // === Valores (siguiendo la lógica del PDF) ===
                        String fecha = safe(reg.getFecha()); // en el payload ya viene formateada
                        String inicioAli = toHoraOrFT(reg.getInicioAlimentacion()); // "FT" si null/empty
                        String finAli    = toHoraOrFT(reg.getFinAlimentacion());    // "FT" si null/empty

                        // Permitidos entero; Tomados/Exceso con 2 decimales (punto)
                        String minPermitidos = reg.getMinutosPermitidos() == null
                                ? "0"
                                : String.valueOf(reg.getMinutosPermitidos().intValue());
                        String minTomados = format2(reg.getMinutosTomados());
                        String minExceso  = format2(reg.getMinutosExceso());

                        // === Escritura de fila ===
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
                        UtilExcel.establecerTexto(r, col++, minPermitidos, null);
                        UtilExcel.establecerTexto(r, col++, minTomados, null);
                        UtilExcel.establecerTexto(r, col++, minExceso, null);
                    }
                }
            }
        }

        int ultimaFila = (filaAct == filaDatosIni) ? filaEnc : (filaAct - 1);

        // 6) Estilos cuerpo
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
            filtros[0] = false; // ITEM sin filtro

            UtilExcel.crearTablaEstilizada(
                    hoja,
                    "TiempoAlimentacionTabla",
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

private String safe(Object v) {
    if (v == null) return "";
    String s = String.valueOf(v).trim();
    return "null".equalsIgnoreCase(s) ? "" : s;
}

// Si viene "yyyy-MM-dd HH:mm:ss" -> devuelve "HH:mm:ss"; si null/"" -> "FT"
private String toHoraOrFT(String fechaHora) {
    if (fechaHora == null || fechaHora.trim().isEmpty()) return "FT";
    int idx = fechaHora.indexOf(' ');
    if (idx < 0 || idx + 1 >= fechaHora.length()) return "FT";
    return fechaHora.substring(idx + 1);
}

private String format2(Double v) {
    if (v == null) return "0.00";
    return String.format("%.2f", v).replace(",", ".");
}


    

}
