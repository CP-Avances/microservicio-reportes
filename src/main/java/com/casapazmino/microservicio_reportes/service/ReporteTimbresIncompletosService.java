package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.apache.poi.ss.usermodel.Row;


import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTimbresIncompletosService {

    public byte[] generarReportePDF(ReporteTimbresIncompletosRequest request) {

        // ➊ DRY: constantes locales (manteniendo look & feel)
        final String TITULO  = "TIMBRES INCOMPLETOS - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
        final String PERIODO = "PERIODO DEL: " + request.getPeriodo().getInicio() + " AL " + request.getPeriodo().getFin();

        final float[] WIDTHS_TITULO = { 8f, 2f };
        final float[] WIDTHS_INFO   = { 4f, 4f, 4f };
        final float[] WIDTHS_TIMBRES = { 1f, 4f, 4f, 4f };

        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_AFTER_TITULO_TABLA = 10f;
        final float SPACING_BEFORE_TIMBRES = 5f;
        final float PADDING_TITULOS = 5f;

        final Color COLOR_PRIMARIO   = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA      = ReporteUtil.colorZebraClaro();

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
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Encabezados
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO));
            document.add(ReporteUtil.crearTituloPeriodo(PERIODO));

            // Contador global (timbres)
            AtomicInteger contadorGlobal = new AtomicInteger(0);
            if (request.getData_pdf() != null) {
                for (TimbresSucursalDTO suc : request.getData_pdf()) {
                    if (suc.getEmpleados() == null) continue;
                    for (EmpleadoDTO emp : suc.getEmpleados()) {
                        if (emp.getTimbres() != null) contadorGlobal.addAndGet(emp.getTimbres().size());
                    }
                }
            }

            // Título + contador
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(WIDTH_PERCENT_100);
            tituloTabla.setWidths(WIDTHS_TITULO);
            tituloTabla.setSpacingAfter(SPACING_AFTER_TITULO_TABLA);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(COLOR_SECUNDARIO);
            celdaTitulo.setPadding(PADDING_TITULOS);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(COLOR_SECUNDARIO);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(PADDING_TITULOS);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            // Iteración sucursales / empleados
            if (request.getData_pdf() != null) {
                for (TimbresSucursalDTO suc : request.getData_pdf()) {
                    if (suc.getEmpleados() == null) continue;

                    for (EmpleadoDTO emp : suc.getEmpleados()) {

                        // 2.1 Bloque info empleado (con borde exterior)
                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                        infoEmpleado.setWidths(WIDTHS_INFO);

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", emp.getApellido() + " " + emp.getNombre(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",     emp.getIdentificacion(),               COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(),               COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:",      emp.getCodigo(),                      COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(),            COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:",        emp.getCargo(),                   COLOR_ZEBRA));

                        PdfPTable contenedorInfo = new PdfPTable(1);
                        contenedorInfo.setWidthPercentage(WIDTH_PERCENT_100);
                        PdfPCell celdaContenedora = new PdfPCell(infoEmpleado);
                        celdaContenedora.setPadding(0);
                        celdaContenedora.setBorder(Rectangle.BOX);
                        contenedorInfo.addCell(celdaContenedora);
                        document.add(contenedorInfo);

                        // 2.2 Tabla de timbres (encabezado con rowspan/colspan)
                        PdfPTable tablaTimbres = new PdfPTable(4);
                        tablaTimbres.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaTimbres.setSpacingBefore(SPACING_BEFORE_TIMBRES);
                        tablaTimbres.setWidths(WIDTHS_TIMBRES);

                        // Fila 1
                        PdfPCell cNum = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        cNum.setRowspan(2);
                        cNum.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cNum.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        cNum.setBorder(Rectangle.BOX);
                        tablaTimbres.addCell(cNum);

                        PdfPCell cTimbre = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        cTimbre.setColspan(2);
                        cTimbre.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cTimbre.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        cTimbre.setBorder(Rectangle.TOP | Rectangle.BOTTOM);
                        tablaTimbres.addCell(cTimbre);

                        PdfPCell cAccion = ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO);
                        cAccion.setRowspan(2);
                        cAccion.setHorizontalAlignment(Element.ALIGN_CENTER);
                        cAccion.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        cAccion.setBorder(Rectangle.BOX);
                        tablaTimbres.addCell(cAccion);

                        // Fila 2
                        tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("HORA",  ReporteUtil.fuenteEncabezado(), COLOR_PRIMARIO));

                        // Cuerpo
                        int contadorLocal = 1;
                        if (emp.getTimbres() != null) {
                            for (TimbreDTO t : emp.getTimbres()) {
                                Color fondo = (contadorLocal % 2 == 0) ? COLOR_ZEBRA : null;

                                String[] partes = t.getFechaHora().split(" ");
                                String fecha = partes.length > 0 ? partes[0] : "";
                                String hora  = partes.length > 1 ? partes[1] : "";

                                tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal), ReporteUtil.fuenteTexto(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(fecha),      ReporteUtil.fuenteTexto(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(hora,                                         ReporteUtil.fuenteTexto(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(traducirAccion(t.getAccion()),                ReporteUtil.fuenteTexto(), fondo));

                                contadorLocal++;
                            }
                        }

                        document.add(tablaTimbres);
                        document.add(Chunk.NEWLINE);
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
            throw new ReportBuildException("No se pudo generar ReporteTimbresIncompletos.pdf", e);
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



    public byte[] generarReporteTimbresIncompletosExcel(ReporteTimbresIncompletosRequest request) {
        System.out.println("Generando XLSX de Timbres Incompletos (una hoja)...");
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            // =========================
            // Hoja única: Timbres Incompletos
            // =========================
            XSSFSheet hoja = libro.createSheet("Timbres Incompletos");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges B1:L5 (B=1 .. L=11 en 0-based)
            for (int row = 0; row <= 4; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, 1, 11);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);
            String activosInactivos = ("1".equals(String.valueOf(request.getOpcionBusqueda()))) ? "ACTIVOS"
                    : "INACTIVOS";
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE TIMBRES INCOMPLETOS - " + activosInactivos, estiloTitulo);
            String periodo = "Periodo del reporte: " +
                    safe(() -> request.getPeriodo().getInicio()) + " al " + safe(() -> request.getPeriodo().getFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEnc = 5;
            String[] headers = {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                    "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                    "FECHA TIMBRE", "HORA TIMBRE", "ACCIÓN"
            };
            int[] anchos = { 10, 20, 18, 28, 18, 18, 18, 20, 18, 20, 16, 26 };

            Row fh = UtilExcel.asegurarFila(hoja, filaEnc);
            for (int c = 0; c < headers.length; c++) {
                UtilExcel.establecerTexto(fh, c, headers[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(fh, headers.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEnc).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado data_pdf → empleados → timbres)
            int filaDatosIni = filaEnc + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getData_pdf() != null) {
                for (TimbresSucursalDTO suc : request.getData_pdf()) {
                    if (suc == null || suc.getEmpleados() == null)
                        continue;

                    for (EmpleadoDTO emp : suc.getEmpleados()) {
                        if (emp == null || emp.getTimbres() == null)
                            continue;

                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                        String ciudad = firstNonEmpty(safe(emp.getCiudad()), safe(suc.getCiudad()));
                        String sucursal = firstNonEmpty(safe(emp.getSucursal()), safe(suc.getSucursal()));

                        for (TimbreDTO t : emp.getTimbres()) {
                            if (t == null)
                                continue;

                            // fechaHora viene "yyyy-MM-dd HH:mm:ss"
                            String fhora = safe(t.getFechaHora());
                            String fecha = "";
                            String hora = safe(t.getHoraTimbre()); // ya viene formateada desde el front
                            if (!fhora.isBlank()) {
                                int i = fhora.indexOf(' ');
                                fecha = i > 0 ? fhora.substring(0, i) : fhora;
                            }
                            String fechaFmt = ReporteUtil.formatearFechaConDia(fecha);
                            String accionTexto = mapAccionTextoLocal(safe(t.getAccion()));


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
                            UtilExcel.establecerTexto(r, col++, fechaFmt, null);
                            UtilExcel.establecerTexto(r, col++, hora, null);
                            UtilExcel.establecerTexto(r, col++, accionTexto, null);
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? filaEnc : (filaAct - 1);

            // 6) Estilos de cuerpo (bordes + alineación)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con bordes
            UtilExcel.aplicarEstiloARegion(hoja, filaEnc, filaEnc, 0, headers.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // Texto largo a la izquierda (apenombre, departamento, cargo)
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 3, 3, estiloIzqBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 7, 8, estiloIzqBorde, true);

                // Resto centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 2, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 4, 6, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 9, 11, estiloCentroBorde, true);

                // 7) Tabla con filtros (ITEM sin filtro)
                boolean[] filtros = new boolean[headers.length];
                for (int i = 0; i < filtros.length; i++)
                    filtros[i] = true;
                filtros[0] = false; // ITEM sin filtro

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TimbresIncompletosReporteTabla",
                        filaEnc, 0,
                        ultimaFila, headers.length - 1,
                        true,
                        filtros);
            }

            // 8) Finalizar
            libro.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /* ===== Helpers locales ===== */
    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String safe(java.util.concurrent.Callable<String> c) {
        try {
            String s = c.call();
            return s == null ? "" : s;
        } catch (Exception e) {
            return "";
        }
    }

    private String firstNonEmpty(String a, String b) {
        return (a == null || a.isBlank()) ? (b == null ? "" : b) : a;
    }

    // Mapeo local para cubrir todos los códigos del Excel antiguo si no llega
    // accionTexto
    private String mapAccionTextoLocal(String cod) {
        if (cod == null)
            return "Desconocido";
        switch (cod.trim().toUpperCase()) {
            case "EOS":
                return "Entrada o salida";
            case "AES":
                return "Inicio o fin alimentación";
            case "PES":
                return "Inicio o fin permiso";
            case "E":
                return "Entrada";
            case "S":
                return "Salida";
            case "I/A":
                return "Inicio alimentación";
            case "F/A":
                return "Fin alimentación";
            case "I/P":
                return "Inicio permiso";
            case "F/P":
                return "Fin permiso";
            case "HA":
                return "Timbre libre";
            default:
                return "Desconocido";
        }
    }


    public static String traducirAccion(String codigo) {
        if (codigo == null)
            return "";

        switch (codigo.trim().toUpperCase()) {
            case "E":
                return "Entrada";
            case "S":
                return "Salida";
            case "I/A":
                return "Inicio alimentación";
            case "F/A":
                return "Fin alimentación";
            default:
                return codigo;
        }
    }

}
