package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.TimbresLibres.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.Collections;
import java.util.List;

@Service
public class ReporteTimbresLibresService {

    // =========================
    // PDF
    // =========================
    public byte[] generarReportePDF(ReporteTimbresLibresRequest request) {
        final boolean conDispositivo = hayColumnaDispositivo(request);
        final String estadoUsuarios = Integer.valueOf(1).equals(request.getOpcionBusqueda()) ? "ACTIVOS" : "INACTIVOS";
        final String TITULO = "TIMBRES ESPECIALES - " + estadoUsuarios;
        final String PERIODO = request.getPeriodo() == null ? "" : "PERIODO DEL: " + safe(request.getPeriodo().getInicio()) + " AL " + safe(request.getPeriodo().getFin());
        final float[] WIDTHS_TITULO = { 8f, 2f };
        final float[] WIDTHS_INFO = { 4f, 4f, 4f };
        final float[] WIDTHS_CON_DISPOSITIVO = { 0.5f, 1.2f, 1.0f, 1.2f, 1.0f, 0.9f, 1.3f, 2.5f, 1.2f, 1.2f };
        final float[] WIDTHS_SIN_DISPOSITIVO = { 0.5f, 1.3f, 1.1f, 0.9f, 1.3f, 3.0f, 1.2f, 1.2f };

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

            document.add(ReporteUtil.crearTituloEmpresa(safe(request.getEmpresa())));
            document.add(ReporteUtil.crearTituloReporte(TITULO));
            if (!PERIODO.isBlank()) document.add(ReporteUtil.crearTituloPeriodo(PERIODO));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

            int totalRegistros = contarRegistrosTotales(request);

            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(WIDTHS_TITULO);
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezadoTablaData()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + totalRegistros, ReporteUtil.fuenteEncabezadoTablaData()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            if (request.getDatos() != null) {
                for (DatoGrupoDTO grupo : request.getDatos()) {
                    if (grupo == null || grupo.getEmpleados() == null) continue;

                    for (EmpleadoDTO emp : grupo.getEmpleados()) {
                        if (emp == null) continue;

                        String nombreCompleto = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();

                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(100);
                        infoEmpleado.setWidths(WIDTHS_INFO);

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", nombreCompleto, zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", safe(emp.getIdentificacion()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", safe(emp.getCodigo()), zebraColor));

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", safe(emp.getRegimen()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", safe(emp.getDepartamento()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", safe(emp.getCargo()), zebraColor));

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CIUDAD:", safe(emp.getCiudad()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("SUCURSAL:", safe(emp.getSucursal()), zebraColor));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("ROL:", safe(emp.getRol()), zebraColor));

                        PdfPTable contenedorInfo = new PdfPTable(1);
                        contenedorInfo.setWidthPercentage(100);

                        PdfPCell celdaContenedora = new PdfPCell(infoEmpleado);
                        celdaContenedora.setPadding(0);
                        celdaContenedora.setBorder(Rectangle.BOX);
                        contenedorInfo.addCell(celdaContenedora);

                        document.add(contenedorInfo);

                        int columnas = conDispositivo ? 10 : 8;
                        PdfPTable tablaTimbres = new PdfPTable(columnas);
                        tablaTimbres.setWidthPercentage(100);
                        tablaTimbres.setSpacingBefore(5f);
                        tablaTimbres.setWidths(conDispositivo ? WIDTHS_CON_DISPOSITIVO : WIDTHS_SIN_DISPOSITIVO);

                        tablaTimbres.addCell(ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));

                        if (conDispositivo) {
                            tablaTimbres.addCell(ReporteUtil.crearCelda("DISPOSITIVO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));
                        }

                        tablaTimbres.addCell(ReporteUtil.crearCelda("RELOJ", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("OBSERVACIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("LONGITUD", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("LATITUD", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));

                        tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
                        tablaTimbres.addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

                        if (conDispositivo) {
                            tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
                            tablaTimbres.addCell(ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
                        }

                        tablaTimbres.setHeaderRows(2);

                        List<TimbreDTO> timbres = emp.getTimbres() == null ? Collections.emptyList() : emp.getTimbres();
                        int contadorLocal = 1;

                        for (TimbreDTO timbre : timbres) {
                            if (timbre == null) continue;

                            Color fondo = contadorLocal % 2 == 0 ? zebraColor : Color.WHITE;

                            tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(formatearFechaSegura(timbre.getFechaServidor()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getHoraServidor()), ReporteUtil.fuenteTablaData(), fondo));

                            if (conDispositivo) {
                                tablaTimbres.addCell(ReporteUtil.crearCelda(formatearFechaSegura(timbre.getFechaDispositivo()), ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getHoraDispositivo()), ReporteUtil.fuenteTablaData(), fondo));
                            }

                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getId_reloj()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(mapAccion(safe(timbre.getAccion())), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getObservacion()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getLongitud()), ReporteUtil.fuenteTablaData(), fondo));
                            tablaTimbres.addCell(ReporteUtil.crearCelda(safe(timbre.getLatitud()), ReporteUtil.fuenteTablaData(), fondo));

                            contadorLocal++;
                        }

                        if (contadorLocal == 1) {
                            PdfPCell sinRegistros = new PdfPCell(new Phrase("SIN REGISTROS", ReporteUtil.fuenteTablaData()));
                            sinRegistros.setColspan(columnas);
                            sinRegistros.setHorizontalAlignment(Element.ALIGN_CENTER);
                            sinRegistros.setVerticalAlignment(Element.ALIGN_MIDDLE);
                            sinRegistros.setBackgroundColor(Color.WHITE);
                            sinRegistros.setPadding(6f);
                            tablaTimbres.addCell(sinRegistros);
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
            throw new ReportBuildException("No se pudo generar ReporteTimbresLibres.pdf", e);
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

    // =========================
    // XLSX (misma estructura ExcelJS)
    // =========================
    public byte[] generarReporteXLSX(ReporteTimbresLibresRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Timbres Especiales"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // ¿Se muestran columnas de fecha/hora del dispositivo?
            final boolean conDispositivo = hayColumnaDispositivo(request);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges de cabecera
            // Si hay dispositivo → B..R (col 1..17); si no → B..P (col 1..15)
            final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
            final int MERGE_COL_INI = 1;
            final int MERGE_COL_FIN = conDispositivo ? 17 : 15;
            for (int fila = MERGE_FIL_INI; fila <= MERGE_FIL_FIN; fila++) {
                UtilExcel.combinarCeldas(hoja, fila, fila, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1,
                    UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);

            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? "LISTA DE TIMBRES ESPECIALES"
                    : request.getTitulo();
            UtilExcel.establecerTexto(hoja, 1, 1,
                    UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);

            if (request.getPeriodo() != null) {
                String periodo = "PERIODO DEL REPORTE: "
                        + safe(request.getPeriodo().getInicio())
                        + " AL "
                        + safe(request.getPeriodo().getFin());
                UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);
            }

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final String[] HEADERS_SIN_DISP = {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                    "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                    "FECHA TIMBRE", "HORA TIMBRE", "RELOJ", "ACCIÓN",
                    "OBSERVACIÓN", "LATITUD", "LONGITUD"
            };
            final int[] ANCHOS_SIN_DISP = {
                    10, 20, 20, 28, 18, 18, 18, 20, 18, 20, 16, 16, 20, 24, 18, 18
            };
            final boolean[] FILTROS_SIN_DISP = {
                    false, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true
            };

            final String[] HEADERS_CON_DISP = {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                    "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                    "FECHA TIMBRE", "HORA TIMBRE", "RELOJ", "ACCIÓN",
                    "OBSERVACIÓN", "LATITUD", "LONGITUD",
                    "FECHA TIMBRE DISPOSITIVO", "HORA TIMBRE DISPOSITIVO"
            };
            final int[] ANCHOS_CON_DISP = {
                    10, 20, 20, 28, 18, 18, 18, 20, 18, 20, 16, 16, 20, 24, 18, 18, 22, 20
            };
            final boolean[] FILTROS_CON_DISP = {
                    false, true, true, true, true, true, true, true, true,
                    true, true, true, true, true, true, true, true, true
            };

            final String[] HEADERS = conDispositivo ? HEADERS_CON_DISP : HEADERS_SIN_DISP;
            final int[] ANCHOS = conDispositivo ? ANCHOS_CON_DISP : ANCHOS_SIN_DISP;
            final boolean[] FILTROS = conDispositivo ? FILTROS_CON_DISP : FILTROS_SIN_DISP;

            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupo → empleado → timbre)
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;
            int item = 1;

            if (request.getDatos() != null) {
                for (DatoGrupoDTO grupo : request.getDatos()) {
                    if (grupo == null || grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoDTO e : grupo.getEmpleados()) {
                        String apenom = (safe(e.getApellido()) + " " + safe(e.getNombre())).trim();
                        List<TimbreDTO> timbres = e.getTimbres();

                        if (timbres == null || timbres.isEmpty()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int c = 0;
                            UtilExcel.establecerValor(r, c++, item++, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCodigo()), null);
                            UtilExcel.establecerTexto(r, c++, apenom, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCiudad()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getSucursal()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getRegimen()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCargo()), null);
                            // columnas de timbre vacías
                            UtilExcel.establecerTexto(r, c++, "", null); // FECHA TIMBRE
                            UtilExcel.establecerTexto(r, c++, "", null); // HORA TIMBRE
                            UtilExcel.establecerTexto(r, c++, "", null); // RELOJ
                            UtilExcel.establecerTexto(r, c++, "", null); // ACCIÓN
                            UtilExcel.establecerTexto(r, c++, "", null); // OBSERVACIÓN
                            UtilExcel.establecerTexto(r, c++, "", null); // LAT
                            UtilExcel.establecerTexto(r, c++, "", null); // LON
                            if (conDispositivo) {
                                UtilExcel.establecerTexto(r, c++, "", null); // FECHA TIMBRE DISPOSITIVO
                                UtilExcel.establecerTexto(r, c++, "", null); // HORA TIMBRE DISPOSITIVO
                            }
                            continue;
                        }

                        for (TimbreDTO t : timbres) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int c = 0;
                            UtilExcel.establecerValor(r, c++, item++, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCodigo()), null);
                            UtilExcel.establecerTexto(r, c++, apenom, null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCiudad()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getSucursal()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getRegimen()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, c++, safe(e.getCargo()), null);

                            UtilExcel.establecerTexto(r, c++, fechaCortaExcel(t.getFechaServidor()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getHoraServidor()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getId_reloj()), null);
                            UtilExcel.establecerTexto(r, c++, mapAccion(safe(t.getAccion())), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getObservacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getLatitud()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getLongitud()), null);
                            if (conDispositivo) {
                                UtilExcel.establecerTexto(r, c++, fechaCortaExcel(t.getFechaDispositivo()), null);
                                UtilExcel.establecerTexto(r, c++, safe(t.getHoraDispositivo()), null);
                            }
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes (por región)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde,
                        true);
            }

            // 7) Tabla estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosIni) {
                String tableName = conDispositivo ? "TimbresReporteTabla" : "TimbresAbiertoReporteTabla";
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        tableName,
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
            throw new ReportBuildException("No se pudo generar TimbresEspeciales.xlsx", e);
        }
    }

    private boolean hayColumnaDispositivo(ReporteTimbresLibresRequest req) {
        if (req == null || req.getDatos() == null)
            return false;
        for (DatoGrupoDTO g : req.getDatos()) {
            if (g.getEmpleados() == null)
                continue;
            for (EmpleadoDTO e : g.getEmpleados()) {
                if (e.getTimbres() == null)
                    continue;
                for (TimbreDTO t : e.getTimbres()) {
                    if ((t.getFechaDispositivo() != null && !t.getFechaDispositivo().trim().isEmpty()) ||
                            (t.getHoraDispositivo() != null && !t.getHoraDispositivo().trim().isEmpty())) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private String mapAccion(String cod) {
        if (cod == null)
            return "";
        String k = cod.trim().toUpperCase();
        switch (k) {
            case "EOS":
                return "Entrada o salida"; // EoS normalizado
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
                return cod; // si ya viene mapeado desde FE, se respeta
        }
    }

    private String safe(String val) {
        return (val == null || val.equalsIgnoreCase("null")) ? "" : val;
    }

    private String fechaCortaExcel(String iso) {
        if (iso == null)
            return "";
        String f = iso.trim();
        return (f.length() >= 10) ? f.substring(0, 10) : f;
    }

    private int contarRegistrosTotales(ReporteTimbresLibresRequest request) {
        int total = 0;
        if (request == null || request.getDatos() == null) return total;

        for (DatoGrupoDTO grupo : request.getDatos()) {
            if (grupo == null || grupo.getEmpleados() == null) continue;

            for (EmpleadoDTO empleado : grupo.getEmpleados()) {
                if (empleado != null && empleado.getTimbres() != null) {
                    total += empleado.getTimbres().size();
                }
            }
        }

        return total;
    }

    private String formatearFechaSegura(String fecha) {
        String valor = safe(fecha);
        if (valor.isBlank()) return "";

        try {
            return ReporteUtil.formatearFechaConDia(valor);
        } catch (Exception e) {
            return valor;
        }
    }
}
