package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

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
        final String estadoUsuarios = Integer.valueOf(1).equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS";
        final String TITULO = "TIMBRES INCOMPLETOS - " + estadoUsuarios;
        final String PERIODO = "PERIODO DEL: " + request.getPeriodo().getInicio() + " AL " + request.getPeriodo().getFin();
        final float[] WIDTHS_TITULO = { 8f, 2f };
        final float[] WIDTHS_INFO = { 4f, 4f, 4f };
        final float[] WIDTHS_TIMBRES = { 0.5f, 1.3f, 1.1f, 1.3f, 1.2f, 1.2f, 1.5f, 1.2f, 1.2f, 1.2f };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 30, 30, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(request.getUsuario(), request.getFraseMarcaAgua(), request.getColorPrincipal()));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO));
            document.add(ReporteUtil.crearTituloPeriodo(PERIODO));
            document.add(ReporteUtil.crearTituloPeriodo("JUSTIFICACIÓN: " + safe(request.getFiltroJustificacion())));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            AtomicInteger contadorGlobal = new AtomicInteger();
            if (request.getData_pdf() != null) {
                for (TimbresSucursalDTO grupo : request.getData_pdf()) {
                    if (grupo == null || grupo.getEmpleados() == null) continue;
                    for (EmpleadoDTO empleado : grupo.getEmpleados()) {
                        if (empleado != null && empleado.getTimbres() != null) contadorGlobal.addAndGet(empleado.getTimbres().size());
                    }
                }
            }

            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(WIDTHS_TITULO);
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezadoTablaData()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezadoTablaData()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);
            document.add(tituloTabla);

            if (request.getData_pdf() != null) {
                for (TimbresSucursalDTO grupo : request.getData_pdf()) {
                    if (grupo == null || grupo.getEmpleados() == null) continue;

                    for (EmpleadoDTO emp : grupo.getEmpleados()) {
                        if (emp == null || emp.getTimbres() == null) continue;

                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(100);
                        infoEmpleado.setWidths(WIDTHS_INFO);
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", safe(emp.getApellido()) + " " + safe(emp.getNombre()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", safe(emp.getIdentificacion()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", safe(emp.getCodigo()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", safe(emp.getRegimen()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", safe(emp.getDepartamento()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", safe(emp.getCargo()), zebraColor));

                        PdfPTable contenedorInfo = new PdfPTable(1);
                        contenedorInfo.setWidthPercentage(100);
                        PdfPCell celdaContenedora = new PdfPCell(infoEmpleado);
                        celdaContenedora.setPadding(0);
                        celdaContenedora.setBorder(Rectangle.BOX);
                        contenedorInfo.addCell(celdaContenedora);
                        document.add(contenedorInfo);

                        PdfPTable tablaTimbres = new PdfPTable(10);
                        tablaTimbres.setWidthPercentage(100);
                        tablaTimbres.setSpacingBefore(5f);
                        tablaTimbres.setWidths(WIDTHS_TIMBRES);

                        tablaTimbres.addCell(ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("HORARIO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("ESTADO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("TIPO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("DESCRIPCIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("DETALLE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                        int contadorLocal = 1;
                        for (TimbreDTO timbre : emp.getTimbres()) {
                            if (timbre == null) continue;

                            Color fondo = contadorLocal % 2 == 0 ? zebraColor : Color.WHITE;
                            String fecha = obtenerFechaHorario(timbre);
                            String hora = obtenerHoraHorario(timbre);

                            tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(formatearFechaSegura(fecha), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(hora, ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(obtenerAccionTexto(timbre), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getEstadoJustificacion()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getTipoJustificacion()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getNombreJustificacion()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(formatearFechaHoraJustificacion(timbre.getDesde()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(formatearFechaHoraJustificacion(timbre.getHasta()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getDetalleJustificacion()), ReporteUtil.fuenteTablaData(), fondo));
                            contadorLocal++;
                        }

                        tablaTimbres.setSpacingAfter(10f);
                        document.add(tablaTimbres);
                    }
                }
            }

            document.close();
            return baos.toByteArray();
        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteTimbresIncompletos.pdf", e);
        } finally {
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
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Timbres Incompletos"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

        // MERGES exactos (B1:L5) → (row 0..4, col 1..11)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 17;

        final String[] HEADERS = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "FECHA HORARIO", "HORA HORARIO", "ACCIÓN", "ESTADO",
                "TIPO", "DESCRIPCIÓN", "DESDE", "HASTA", "DETALLE"
        };

        final int[] ANCHOS = {
                10, 20, 18, 28,
                18, 20, 20, 25, 22,
                20, 18, 25, 20,
                20, 25, 25, 25, 20
        };

        final boolean[] FILTROS = {
                false, true, true, true,
                true, true, true, true, true,
                true, true, true, true,
                true, true, true, true, true
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
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES B1:L5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1,
                    UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())), estiloTitulo);

            String activosInactivos = ("1".equals(String.valueOf(request.getOpcionBusqueda())))
                    ? "ACTIVOS"
                    : "INACTIVOS";
            UtilExcel.establecerTexto(hoja, 1, 1,
                    "LISTA DE TIMBRES INCOMPLETOS - " + activosInactivos, estiloTitulo);

            String periodo = "Periodo del reporte: "
                    + safe(() -> request.getPeriodo().getInicio())
                    + " al "
                    + safe(() -> request.getPeriodo().getFin());
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

            UtilExcel.establecerTexto(hoja, 3, 1,
                "FILTRO DE JUSTIFICACIÓN: " + safe(request.getFiltroJustificacion()),
                estiloTitulo);

            // 4) ENCABEZADOS + ANCHOS (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO (aplanado data_pdf → empleados → timbres)
            int filaDatosIni = FILA_ENCABEZADO + 1;
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
                            if (t == null) continue;

                            String fechaFmt = formatearFechaSegura(obtenerFechaHorario(t));
                            String hora = obtenerHoraHorario(t);
                            String accion = obtenerAccionTexto(t);
                            String estado = safe(t.getEstadoJustificacion());
                            String tipo = safe(t.getTipoJustificacion());
                            String descripcion = safe(t.getNombreJustificacion());
                            String desde = formatearFechaHoraJustificacion(t.getDesde());
                            String hasta = formatearFechaHoraJustificacion(t.getHasta());
                            String detalle = safe(t.getDetalleJustificacion());

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
                            UtilExcel.establecerTexto(r, col++, accion, null);
                            UtilExcel.establecerTexto(r, col++, estado, null);
                            UtilExcel.establecerTexto(r, col++, tipo, null);
                            UtilExcel.establecerTexto(r, col++, descripcion, null);
                            UtilExcel.establecerTexto(r, col++, desde, null);
                            UtilExcel.establecerTexto(r, col++, hasta, null);
                            UtilExcel.establecerTexto(r, col++, detalle, null);
                        }

                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) ESTILOS de cuerpo (manteniendo look & feel original)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con bordes
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, HEADERS.length - 1, estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 3, 3, estiloIzqBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 7, 8, estiloIzqBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 14, 14, estiloIzqBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 17, 17, estiloIzqBorde, true);

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "TimbresIncompletosReporteTabla",
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
            throw new ReportBuildException("No se pudo generar TimbresIncompletos.xlsx", e);
        }
    }

    private String obtenerFechaHorario(TimbreDTO timbre) {
        String fecha = safe(timbre.getFechaHorario());
        if (!fecha.isBlank()) return fecha;

        String fechaHora = safe(timbre.getFechaHora());
        int separador = fechaHora.indexOf(' ');
        return separador > 0 ? fechaHora.substring(0, separador) : fechaHora;
    }

    private String obtenerHoraHorario(TimbreDTO timbre) {
        String hora = firstNonEmpty(safe(timbre.getHoraHorario()), safe(timbre.getHoraTimbre()));
        if (!hora.isBlank()) return hora;

        String fechaHora = safe(timbre.getFechaHora());
        int separador = fechaHora.indexOf(' ');
        return separador >= 0 && separador + 1 < fechaHora.length()
                ? fechaHora.substring(separador + 1)
                : "";
    }

    private String obtenerAccionTexto(TimbreDTO timbre) {
        return firstNonEmpty(safe(timbre.getAccionTexto()), mapAccionTextoLocal(safe(timbre.getAccion())));
    }

    private String formatearFechaSegura(String fecha) {
        String valor = safe(fecha);
        return valor.isBlank() ? "" : ReporteUtil.formatearFechaConDia(valor);
    }

    private String formatearFechaHoraJustificacion(String valor) {
        String texto = safe(valor);
        if (texto.isBlank()) return "";

        String[] partes = texto.split("\\s+");
        String fecha = formatearFechaSegura(partes[0]);

        if (partes.length < 2) return fecha;

        StringBuilder hora = new StringBuilder();
        for (int i = 1; i < partes.length; i++) {
            if (i > 1) hora.append(" ");
            hora.append(partes[i]);
        }

        return fecha + " " + hora;
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
