package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales.*;
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
public class ReporteTimbresVirtualesService {

    public byte[] generarReportePDF(ReporteTimbresVirtualesRequest request) {
        // DRY: constantes locales
        final float[] WIDTHS_TITULO_TABLA   = { 8f, 2f };
        final float[] WIDTHS_INFO_EMPLEADO  = { 4f, 4f, 4f };
        final float[] WIDTHS_TABLA_TIMBRES  = { 1f, 2.5f, 1.5f, 1.5f, 2.3f, 4.5f, 2f, 2f };
        final String  TITULO_REPORTE        = "TIMBRES VIRTUALES - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
        final String  TITULO_PERIODO        = "PERIODO DEL: " + request.getPeriodo().getInicio() + " AL " + request.getPeriodo().getFin();
        final String  TITULO_LISTA          = "LISTA EMPLEADOS";

        // Colores calculados una sola vez
        final Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color zebraColor      = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer  = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
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

            // 2) Construcción (helpers existentes, sin cambiar look & feel)
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
                g -> g.getEmpleados().forEach(e -> contadorGlobal.addAndGet(e.getTimbres().size()))
            );

            // Banda de título + contador
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

            // Iteración por grupos y empleados
            for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                for (EmpleadoTimbreDTO emp : grupo.getEmpleados()) {

                    // Info empleado
                    PdfPTable infoEmpleado = new PdfPTable(3);
                    infoEmpleado.setWidthPercentage(100);
                    infoEmpleado.setWidths(WIDTHS_INFO_EMPLEADO);

                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", emp.getApellido() + " " + emp.getNombre(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
                    infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

                    PdfPTable tablaContenedora = new PdfPTable(1);
                    tablaContenedora.setWidthPercentage(100);
                    PdfPCell contenedor = new PdfPCell(infoEmpleado);
                    contenedor.setPadding(0);
                    contenedor.setBorder(Rectangle.BOX);
                    tablaContenedora.addCell(contenedor);
                    document.add(tablaContenedora);

                    // Tabla de timbres
                    PdfPTable tablaTimbres = new PdfPTable(8);
                    tablaTimbres.setWidthPercentage(100);
                    tablaTimbres.setSpacingBefore(5f);
                    tablaTimbres.setWidths(WIDTHS_TABLA_TIMBRES);

                    // Encabezado fila 1
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

                    // Encabezado fila 2
                    tablaTimbres.addCell(ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezado(), colorPrincipal));
                    tablaTimbres.addCell(ReporteUtil.crearCelda("HORA",  ReporteUtil.fuenteEncabezado(), colorPrincipal));

                    // Cuerpo
                    int contadorLocal = 1;
                    for (TimbreUsuarioDTO t : emp.getTimbres()) {
                        Color fondo = (contadorLocal % 2 == 0) ? zebraColor : Color.WHITE;

                        String[] partes = t.getFecha_hora_timbre().split(" ");
                        String fecha = (partes.length > 0) ? partes[0] : "";
                        String hora  = (partes.length > 1) ? partes[1] : "";

                        tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal),            ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(fecha),  ReporteUtil.fuenteTexto(), fondo));
                        tablaTimbres.addCell(ReporteUtil.crearCelda(hora,                                    ReporteUtil.fuenteTexto(), fondo));
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
            // Dejar que el controller mapee a 400 si corresponde
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar TimbresVirtuales.pdf", e);
        } finally {
            // Ciclo de recursos garantizado
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


    public byte[] generarReporteXLSX(ReporteTimbresVirtualesRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Timbres Virtuales");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            final boolean conDispositivo = request.getTimbreDispositivo();

            // 2) Merges cabecera (5 filas)
            // Con dispositivo -> B..R (indices 1..17)
            // Sin dispositivo -> B..P (indices 1..15)
            if (conDispositivo) {
                for (int fila = 0; fila <= 4; fila++)
                    UtilExcel.combinarCeldas(hoja, fila, fila, 1, 17);
            } else {
                for (int fila = 0; fila <= 4; fila++)
                    UtilExcel.combinarCeldas(hoja, fila, fila, 1, 15);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            String titulo = "LISTA DE TIMBRES VIRTUALES";
            UtilExcel.establecerTexto(hoja, 1, 1, UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);
            if (request.getPeriodo() != null) {
                UtilExcel.establecerTexto(hoja, 2, 1,
                        "PERIODO DEL REPORTE: " + safe(request.getPeriodo().getInicio()) + " AL "
                                + safe(request.getPeriodo().getFin()),
                        null);
            }

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;

            String[] headersSinDisp = new String[] {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                    "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                    "SERVIDOR FECHA", "SERVIDOR HORA", "ID RELOJ", "ACCIÓN", "OBSERVACIÓN", "LATITUD", "LONGITUD"
            };
            int[] anchosSinDisp = new int[] { 10, 20, 20, 22, 18, 18, 18, 20, 18, 18, 14, 14, 18, 24, 14, 14 };

            String[] headersConDisp = new String[] {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE",
                    "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                    "SERVIDOR FECHA", "SERVIDOR HORA", "ID RELOJ", "ACCIÓN", "OBSERVACIÓN", "LATITUD", "LONGITUD",
                    "FECHA TIMBRE DISPOSITIVO", "HORA TIMBRE DISPOSITIVO"
            };
            int[] anchosConDisp = new int[] { 10, 20, 20, 22, 18, 18, 18, 20, 18, 18, 14, 14, 18, 24, 14, 14, 20, 18 };

            String[] headers = conDispositivo ? headersConDisp : headersSinDisp;
            int[] anchos = conDispositivo ? anchosConDisp : anchosSinDisp;

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < headers.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, headers[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, headers.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupo → empleado → timbre)
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            if (request.getData_pdf() != null) {
                for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoTimbreDTO usu : grupo.getEmpleados()) {
                        String apenom = (safe(usu.getApellido()) + " " + safe(usu.getNombre())).trim();

                        if (usu.getTimbres() == null || usu.getTimbres().isEmpty()) {
                            // fila vacía si no hay timbres
                            Row r = UtilExcel.asegurarFila(hoja, filaActual++);
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
                            UtilExcel.establecerTexto(r, c++, "", null); // accion
                            UtilExcel.establecerTexto(r, c++, "", null); // obs
                            UtilExcel.establecerTexto(r, c++, "", null); // lat
                            UtilExcel.establecerTexto(r, c++, "", null); // lon
                            if (conDispositivo) {
                                UtilExcel.establecerTexto(r, c++, "", null); // fecha disp
                                UtilExcel.establecerTexto(r, c++, "", null); // hora disp
                            }
                            continue;
                        }

                        for (TimbreUsuarioDTO t : usu.getTimbres()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaActual++);
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

                            // servidor fecha/hora desde "fecha_hora_timbre_validado"
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

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, headers.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, headers.length - 1, estiloIzqBorde,
                        true);
            }

            // 7) Tabla estilizada (zebra + AutoFilter)
            String tableName = conDispositivo ? "TimbresVirtualesReporteTabla" : "TimbresVirtualesReporteTabla";
            if (ultimaFila >= filaDatosInicio) {
                boolean[] filtros = new boolean[headers.length];
                for (int i = 0; i < filtros.length; i++)
                    filtros[i] = true;
                filtros[0] = false; // ITEM sin filtro

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        tableName,
                        filaEncabezado, 0,
                        ultimaFila, headers.length - 1,
                        true,
                        filtros);
            }

            libro.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    /* ====================== Helpers locales ====================== */
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
