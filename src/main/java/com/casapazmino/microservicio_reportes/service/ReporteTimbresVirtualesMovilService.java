package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtualesMovil.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.Element;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.Rectangle;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

// Apache POI
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTimbresVirtualesMovilService {

    public byte[] generarReportePDF(ReporteTimbresVirtualesMovilRequest request) {
        // DRY: constantes locales
        final float[] WIDTHS_TITULO_TABLA   = { 8f, 2f };
        final float[] WIDTHS_INFO_EMPLEADO  = { 4f, 4f, 4f };
        final float[] WIDTHS_TABLA_TIMBRES  = { 1f, 2.5f, 1.5f, 1.5f, 2.3f, 4.5f, 2f, 2f };
        final String  TITULO_LISTA          = "LISTA EMPLEADOS";
        final String  TITULO_REPORTE        = "TIMBRES (APLICACIÓN MOVIL) - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
        final String  TITULO_PERIODO        = "PERIODO DEL: " + request.getPeriodo().getInicio() + " AL " + request.getPeriodo().getFin();

        // Colores calculados una vez
        final Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color zebraColor      = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer  = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización de recursos PDF
            Rectangle orientacion = request.getTimbreDispositivo() ? PageSize.A4.rotate() : PageSize.A4;
            baos     = new ByteArrayOutputStream();
            document = new Document(orientacion, 40, 40, 30, 50);
            writer   = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (solo helpers existentes; no cambiamos diseño)
            // Logo y encabezados
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO_REPORTE));
            document.add(ReporteUtil.crearTituloPeriodo(TITULO_PERIODO));

            // Contador global de timbres
            AtomicInteger contadorGlobal = new AtomicInteger();
            request.getData_pdf().forEach(
                grupo -> grupo.getEmpleados().forEach(emp -> contadorGlobal.addAndGet(emp.getTimbres().size()))
            );

            // Título tabla (LISTA EMPLEADOS + contador)
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(WIDTHS_TITULO_TABLA);
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase(TITULO_LISTA, ReporteUtil.fuenteEncabezado()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezado()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            // Cuerpo por grupos y empleados
            for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                for (EmpleadoTimbreDTO emp : grupo.getEmpleados()) {

                    // Bloque info empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(WIDTHS_INFO_EMPLEADO);

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",  emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:",      emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:",       emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:",     emp.getCargo(), zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    document.add(tablaContenedora);

                    // Tabla timbres
                    PdfPTable tablaTimbres = new PdfPTable(8);
                    tablaTimbres.setWidthPercentage(100);
                    tablaTimbres.setSpacingBefore(5f);
                    tablaTimbres.setWidths(WIDTHS_TABLA_TIMBRES);

                    // Encabezado Fila 1
                    PdfPCell celdaN = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaN.setRowspan(2);
                    celdaN.setHorizontalAlignment(Element.ALIGN_CENTER);
                    tablaTimbres.addCell(celdaN);

                    PdfPCell celdaTimbre = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaTimbre.setColspan(2);
                    celdaTimbre.setHorizontalAlignment(Element.ALIGN_CENTER);
                    tablaTimbres.addCell(celdaTimbre);

                    PdfPCell celdaReloj = ReporteUtil.crearCelda("RELOJ", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaReloj.setRowspan(2);
                    tablaTimbres.addCell(celdaReloj);

                    PdfPCell celdaAccion = ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaAccion.setRowspan(2);
                    tablaTimbres.addCell(celdaAccion);

                    PdfPCell celdaObs = ReporteUtil.crearCelda("OBSERVACIÓN", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaObs.setRowspan(2);
                    tablaTimbres.addCell(celdaObs);

                    PdfPCell celdaLong = ReporteUtil.crearCelda("LONGITUD", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaLong.setRowspan(2);
                    tablaTimbres.addCell(celdaLong);

                    PdfPCell celdaLat = ReporteUtil.crearCelda("LATITUD", ReporteUtil.fuenteEncabezado(), colorPrincipal);
                    celdaLat.setRowspan(2);
                    tablaTimbres.addCell(celdaLat);

                    // Encabezado Fila 2
                    tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaTimbres.addCell(ReporteUtil.crearCelda("HORA",  ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    // Cuerpo
                    int contadorLocal = 1;
                    for (TimbreUsuarioDTO t : emp.getTimbres()) {
                        Color fondo = (contadorLocal % 2 == 0) ? zebraColor : Color.WHITE;

                        String[] partes = t.getFecha_hora_timbre().split(" ");
                        String fecha = (partes.length > 0) ? partes[0] : "";
                        String hora  = (partes.length > 1) ? partes[1] : "";

                        tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal),           ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(fecha), ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(hora,                                   ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getId_reloj(),                         ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.traducirAccion(t.getAccion()),ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getObservacion(),                      ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getLongitud(),                         ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(t.getLatitud(),                          ReporteUtil.fuenteTexto(), fondo));

                        contadorLocal++;
                    }

                    document.add(tablaTimbres);
                    document.add(Chunk.NEWLINE);
                }
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → que el controller decida si es 400
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar TimbresVirtualesMovil.pdf", e);
        } finally {
            // Ciclo de recursos garantizado (cerrar silenciosamente)
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
    // XLSX
    // =========================
    public byte[] generarReporteXLSX(ReporteTimbresVirtualesMovilRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA     = "Timbres Virtuales Movil"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;                         // fila 6 (idx 5)

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            final boolean conDispositivo = request.getTimbreDispositivo();

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) Merges cabecera (5 filas): con disp → B..R (1..17) | sin disp → B..P (1..15)
            final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
            final int MERGE_COL_INI = 1;
            final int MERGE_COL_FIN = conDispositivo ? 17 : 15;
            for (int fila = MERGE_FIL_INI; fila <= MERGE_FIL_FIN; fila++) {
                UtilExcel.combinarCeldas(hoja, fila, fila, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, UtilExcel.aMayusculasSeguras("LISTA DE TIMBRES VIRTUALES MOVIL"), estiloTitulo);
            if (request.getPeriodo() != null) {
                String periodo = "PERIODO DEL REPORTE: " + safe(request.getPeriodo().getInicio()) +
                                " AL " + safe(request.getPeriodo().getFin());
                UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);
            }

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final String[] HEADERS_SIN_DISP = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "SERVIDOR FECHA", "SERVIDOR HORA", "ID RELOJ", "ACCIÓN",
                "OBSERVACIÓN", "LATITUD", "LONGITUD"
            };
            final int[] ANCHOS_SIN_DISP = { 10, 20, 20, 22, 18, 18, 18, 20, 18, 20, 16, 16, 18, 26, 16, 16 };
            final boolean[] FILTROS_SIN_DISP = {
                false, true, true, true, true, true, true, true, true,
                true,  true, true, true, true, true, true
            };

            final String[] HEADERS_CON_DISP = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                "SERVIDOR FECHA", "SERVIDOR HORA", "ID RELOJ", "ACCIÓN",
                "OBSERVACIÓN", "LATITUD", "LONGITUD",
                "FECHA TIMBRE DISPOSITIVO", "HORA TIMBRE DISPOSITIVO"
            };
            final int[] ANCHOS_CON_DISP = { 10, 20, 20, 22, 18, 18, 18, 20, 18, 20, 16, 16, 18, 26, 16, 16, 20, 20 };
            final boolean[] FILTROS_CON_DISP = {
                false, true, true, true, true, true, true, true, true,
                true,  true, true, true, true, true, true, true, true
            };

            final String[] HEADERS  = conDispositivo ? HEADERS_CON_DISP  : HEADERS_SIN_DISP;
            final int[]    ANCHOS   = conDispositivo ? ANCHOS_CON_DISP   : ANCHOS_SIN_DISP;
            final boolean[] FILTROS = conDispositivo ? FILTROS_CON_DISP  : FILTROS_SIN_DISP;

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

            if (request.getData_pdf() != null) {
                for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                    if (grupo == null || grupo.getEmpleados() == null) continue;

                    for (EmpleadoTimbreDTO usu : grupo.getEmpleados()) {
                        String apenom = (safe(usu.getApellido()) + " " + safe(usu.getNombre())).trim();

                        if (usu.getTimbres() == null || usu.getTimbres().isEmpty()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int c = 0;
                            UtilExcel.establecerValor(r, c++, item++, null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getCodigo()), null);
                            UtilExcel.establecerTexto(r, c++, apenom, null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getCiudad()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getSucursal()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getRegimen()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getCargo()), null);
                            UtilExcel.establecerTexto(r, c++, "", null); // servidor fecha
                            UtilExcel.establecerTexto(r, c++, "", null); // servidor hora
                            UtilExcel.establecerTexto(r, c++, "", null); // id reloj
                            UtilExcel.establecerTexto(r, c++, "", null); // acción
                            UtilExcel.establecerTexto(r, c++, "", null); // observación
                            UtilExcel.establecerTexto(r, c++, "", null); // latitud
                            UtilExcel.establecerTexto(r, c++, "", null); // longitud
                            if (conDispositivo) {
                                UtilExcel.establecerTexto(r, c++, "", null); // fecha disp
                                UtilExcel.establecerTexto(r, c++, "", null); // hora disp
                            }
                            continue;
                        }

                        for (TimbreUsuarioDTO t : usu.getTimbres()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                            int c = 0;

                            UtilExcel.establecerValor(r, c++, item++, null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getCodigo()), null);
                            UtilExcel.establecerTexto(r, c++, apenom, null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getCiudad()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getSucursal()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getRegimen()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, c++, safe(usu.getCargo()), null);

                            // Servidor fecha/hora desde "fecha_hora_timbre_validado"
                            String fhServ = safe(t.getFecha_hora_timbre_validado());
                            UtilExcel.establecerTexto(r, c++, fechaCortaExcel(fhServ), null);
                            UtilExcel.establecerTexto(r, c++, extraerHora(fhServ), null);

                            UtilExcel.establecerTexto(r, c++, safe(t.getId_reloj()), null);
                            UtilExcel.establecerTexto(r, c++, ReporteUtil.traducirAccion(safe(t.getAccion())), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getObservacion()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getLatitud()), null);
                            UtilExcel.establecerTexto(r, c++, safe(t.getLongitud()), null);

                            if (conDispositivo) {
                                String fhDisp = safe(t.getFecha_hora_timbre());
                                UtilExcel.establecerTexto(r, c++, fechaCortaExcel(fhDisp), null);
                                UtilExcel.establecerTexto(r, c++, extraerHora(fhDisp), null);
                            }
                        }
                    }
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineaciones + bordes (por región)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con bordes
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            if (ultimaFila >= filaDatosIni) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde, true);
            }

            // 7) Tabla estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "TimbresMovilReporteTabla",
                    FILA_ENCABEZADO, 0,
                    ultimaFila, HEADERS.length - 1,
                    true,
                    FILTROS
                );
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar TimbresVirtualesMovil.xlsx", e);
        }
    }

    
    /* ===== Helpers locales ===== */
    private String safe(String v) {
        return (v == null || v.equalsIgnoreCase("null")) ? "" : v;
    }

    private String fechaCortaExcel(String iso) {
        if (iso == null)
            return "";
        String f = iso.trim();
        return (f.length() >= 10) ? f.substring(0, 10) : f;
    }

    private String extraerHora(String iso) {
        if (iso == null)
            return "";
        String f = iso.trim();
        int sp = f.indexOf(' ');
        return (sp > 0 && sp + 1 < f.length()) ? f.substring(sp + 1) : "";
    }

}
