package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion.*;
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
public class ReporteTiempoAlimentacionService {

    public byte[] generarReporteTiempoAlimentacionPDF(ReporteTiempoAlimentacionRequest request) {

        // ➊ DRY: constantes locales (look & feel intacto)
        final String TITULO = "TIEMPO DE ALIMENTACIÓN - " + ("1".equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS");
        final String PERIODO = "PERIODO DEL: " + request.getFechaInicio() + " AL " + request.getFechaFin();

        final float[] WIDTHS_COLORES = { 3f, 1.5f, 1.5f, 2f, 2f };
        final float[] WIDTHS_TITULO  = { 8f, 2f };
        final float[] WIDTHS_INFO    = { 4f, 4f, 4f };
        final float[] WIDTHS_TABLA   = { 0.5f, 1.8f, 1.8f, 1.8f, 1.8f, 1.8f, 1.8f };

        final String[] ENCAB_COLORES = { "CÓDIGO DE COLOR", "FALTA TIMBRE", " ", "EXCESO DE ALIMENTACIÓN", " " };
        final String[] ENCAB_TABLA   = { "Nº", "FECHA", "INICIO ALIMENTACIÓN", "FIN ALIMENTACIÓN", "M. ALIMENTACIÓN",
                                        "M. TOMADOS", "M. EXCESO" };

        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_AFTER_BLOQUE = 10f;
        final float SPACING_AFTER_CONTENEDOR = 3f;
        final float PADDING_TITULOS = 5f;

        final Color COLOR_PRIMARIO   = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA      = ReporteUtil.colorZebraClaro();
        final Color COLOR_EXCESO     = new Color(0x55EE44);
        final Color COLOR_FT         = new Color(0xEE4444);

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4, 40, 40, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

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

            colores.setSpacingAfter(SPACING_AFTER_BLOQUE);
            document.add(colores);

            // Total de registros (minas)
            AtomicInteger contadorGlobal = new AtomicInteger();
            if (request.getGrupos() != null) {
                request.getGrupos().forEach(g -> {
                    if (g.getEmpleados() != null) {
                        g.getEmpleados().forEach(e ->
                            contadorGlobal.addAndGet(e.getAlimentacion() != null ? e.getAlimentacion().size() : 0)
                        );
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

            PdfPCell celda2 = new PdfPCell(new Phrase("Nº Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
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
                    if (grupo.getEmpleados() == null) continue;

                    for (EmpleadoAlimentacionDTO emp : grupo.getEmpleados()) {

                        double totalExcesoAlimentacion = 0d;
                        int contador = 1;

                        // Encabezado info empleado con borde exterior
                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                        infoEmpleado.setWidths(WIDTHS_INFO);

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",  emp.getApellido() + " " + emp.getNombre(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",      emp.getIdentificacion(),                COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:",      emp.getCodigo(),                        COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(),                COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:",    emp.getDepartamento(),           COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:",           emp.getCargo(),                  COLOR_ZEBRA));

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
                            tablaAlimentacion.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        }

                        if (emp.getAlimentacion() != null) {
                            for (RegistroAlimentacionDTO registro : emp.getAlimentacion()) {
                                Color fondo = (contador % 2 == 0) ? COLOR_ZEBRA : Color.WHITE;

                                String inicio = (registro.getInicioAlimentacion() == null || registro.getInicioAlimentacion().isEmpty())
                                                ? "FT" : extraerHora(registro.getInicioAlimentacion());
                                String fin    = (registro.getFinAlimentacion() == null || registro.getFinAlimentacion().isEmpty())
                                                ? "FT" : extraerHora(registro.getFinAlimentacion());

                                Color fondoInicio = "FT".equals(inicio) ? COLOR_FT : fondo;
                                Color fondoFin    = "FT".equals(fin)    ? COLOR_FT : fondo;

                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(String.valueOf(contador), fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(registro.getFecha(), fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(inicio, fondoInicio));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(fin,   fondoFin));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(String.valueOf(registro.getMinutosPermitidos()), fondo));
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(String.valueOf(registro.getMinutosTomados()).replace(",", "."), fondo));

                                Color fondoExceso = registro.getMinutosExceso() > 0 ? COLOR_EXCESO : fondo;
                                tablaAlimentacion.addCell(ReporteUtil.celdaCentro(
                                        String.format("%.2f", registro.getMinutosExceso()).replace(",", "."),
                                        fondoExceso
                                ));

                                totalExcesoAlimentacion += registro.getMinutosExceso();
                                contador++;
                            }
                        }

                        // Fila de totales: 5 vacías + "TOTAL" + total exceso
                        for (int i = 0; i < 5; i++) {
                            PdfPCell celdaVacia = ReporteUtil.crearCelda("", ReporteUtil.fuenteTexto(), Color.WHITE);
                            celdaVacia.setBorder(Rectangle.NO_BORDER);
                            tablaAlimentacion.addCell(celdaVacia);
                        }
                        tablaAlimentacion.addCell(ReporteUtil.crearCelda("TOTAL", ReporteUtil.fuenteTexto(), Color.WHITE));
                        tablaAlimentacion.addCell(ReporteUtil.crearCelda(
                                String.format("%.2f", totalExcesoAlimentacion).replace(",", "."),
                                ReporteUtil.fuenteTexto(), Color.WHITE));

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
