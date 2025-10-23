package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteVacunacion.AgrupadorVacunaUsuarioDTO;
import com.casapazmino.microservicio_reportes.model.ReporteVacunacion.EmpleadoVacunaUsuarioDTO;
import com.casapazmino.microservicio_reportes.model.ReporteVacunacion.ReporteVacunacionUsuariosRequest;
import com.casapazmino.microservicio_reportes.model.ReporteVacunacion.VacunaUsuarioDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

@Service
public class ReporteVacunacionUsuariosService {

    public byte[] generarReportePDF(ReporteVacunacionUsuariosRequest request) {
        // DRY: constantes locales
        final float   M_I = 40f, M_D = 40f, M_S = 50f, M_INF = 50f;
        final float[] WIDTHS_CABECERA   = { 3f, 3f, 2f };
        final float[] WIDTHS_EMPLEADO   = { 3f, 4f, 3f };
        final float[] WIDTHS_VACUNAS    = { 1f, 3f, 2f, 5f };
        final String  PREF_SUCURSAL     = "SUCURSAL: ";
        final String  PREF_REGISTROS    = "N° Registros: ";

        Document document = null;
        PdfWriter writer  = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización de recursos PDF
            baos     = new ByteArrayOutputStream();
            document = new Document(PageSize.A4, M_I, M_D, M_S, M_INF);
            writer   = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));
            document.open();

            // 2) Construcción (solo helpers existentes; no cambiamos diseño)
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            // Encabezados
            Paragraph empresa = new Paragraph(request.getEmpresa(), ReporteUtil.fuenteEncabezado());
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            Paragraph titulo = new Paragraph(safe(request.getTitulo()), ReporteUtil.fuenteEncabezado());
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            // Colores y fuente
            final Color colorPrincipal  = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            final Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            final Font  fuente          = ReporteUtil.fuenteTexto();

            for (AgrupadorVacunaUsuarioDTO grupo : request.getDatos()) {

                // Descripción por tipo de filtro
                String descripcion = "";
                String establecimiento = safe(PREF_SUCURSAL + grupo.getSucursal());

                switch (safe(request.getTipoFiltro()).toLowerCase()) {
                    case "regimen":
                        descripcion = "REGIMEN: " + safe(grupo.getNombre());
                        break;
                    case "departamento":
                        descripcion = "DEPARTAMENTO: " + safe(grupo.getDepartamento());
                        break;
                    case "cargo":
                        descripcion = "CARGO: " + safe(grupo.getNombre());
                        break;
                    case "ciudad":
                        descripcion = "CIUDAD: " + safe(grupo.getCiudad());
                        break;
                    case "empleado":
                        descripcion = "LISTA EMPLEADOS";
                        establecimiento = "";
                        break;
                }

                int totalRegistros = grupo.getEmpleados().stream()
                        .mapToInt(e -> e.getVacunas().size())
                        .sum();

                // Cabecera de bloque (verde, sin bordes internos; borde externo con TableEvent)
                PdfPTable tablaCabecera = new PdfPTable(3);
                tablaCabecera.setWidthPercentage(100);
                tablaCabecera.setWidths(WIDTHS_CABECERA);
                tablaCabecera.setSpacingBefore(10f);
                tablaCabecera.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                tablaCabecera.addCell(celdaSinBordeIzquierda(descripcion,     fuente, colorSecundario));
                tablaCabecera.addCell(celdaSinBordeIzquierda(establecimiento,  fuente, colorSecundario));
                tablaCabecera.addCell(celdaSinBordeIzquierda(PREF_REGISTROS + totalRegistros, fuente, colorSecundario));

                tablaCabecera.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                    PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                    cb.rectangle(
                            widths[0][0],
                            heights[heights.length - 1],
                            widths[0][widths[0].length - 1] - widths[0][0],
                            heights[0] - heights[heights.length - 1]
                    );
                    cb.stroke();
                });
                document.add(tablaCabecera);

                // Empleado + detalle de vacunas
                for (EmpleadoVacunaUsuarioDTO empl : grupo.getEmpleados()) {
                    // Tabla de información del empleado (3 columnas, gris claro, borde externo)
                    PdfPTable tablaEmpleado = new PdfPTable(3);
                    tablaEmpleado.setWidthPercentage(100);
                    tablaEmpleado.setSpacingBefore(6f);
                    tablaEmpleado.setWidths(WIDTHS_EMPLEADO);

                    String[][] filas = new String[][] {
                        { "C.C.: " + safe(empl.getIdentificacion()),
                        "EMPLEADO: " + safe(empl.getApellido()) + " " + safe(empl.getNombre()),
                        "DEPARTAMENTO: " + safe(empl.getDepartamento()) },
                        { "CORREO: " + safe(empl.getCorreo()),
                        "GENERO: " + safe(empl.getGenero()),
                        "CARGO: " + safe(empl.getCargo()) },
                        { "REGIMEN: " + safe(empl.getRegimen()),
                        "COD: " + safe(empl.getCodigo()),
                        "ROL: " + safe(empl.getRol()) }
                    };

                    for (String[] fila : filas) {
                        for (String texto : fila) {
                            PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
                            celda.setBackgroundColor(Color.LIGHT_GRAY);
                            celda.setHorizontalAlignment(Element.ALIGN_LEFT);
                            celda.setBorder(Rectangle.NO_BORDER);
                            tablaEmpleado.addCell(celda);
                        }
                    }

                    tablaEmpleado.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                        PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                        cb.rectangle(
                                widths[0][0],
                                heights[heights.length - 1],
                                widths[0][widths[0].length - 1] - widths[0][0],
                                heights[0] - heights[heights.length - 1]
                        );
                        cb.stroke();
                    });
                    document.add(tablaEmpleado);

                    // Tabla de vacunas (4 columnas)
                    PdfPTable tablaVacunas = new PdfPTable(4);
                    tablaVacunas.setWidthPercentage(100);
                    tablaVacunas.setSpacingBefore(0f);
                    tablaVacunas.setWidths(WIDTHS_VACUNAS);

                    // Encabezados centrados
                    for (String header : new String[] { "N°", "VACUNA", "FECHA", "DESCRIPCIÓN" }) {
                        PdfPCell headerCell = new PdfPCell(new Phrase(header, fuente));
                        headerCell.setBackgroundColor(colorPrincipal);
                        headerCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        tablaVacunas.addCell(headerCell);
                    }

                    int index = 1;
                    for (VacunaUsuarioDTO vac : empl.getVacunas()) {
                        PdfPCell c1 = new PdfPCell(new Phrase(String.valueOf(index++),           fuente));
                        PdfPCell c2 = new PdfPCell(new Phrase(safe(vac.getTipo_vacuna()),        fuente));
                        PdfPCell c3 = new PdfPCell(new Phrase(formatearFecha(vac.getFecha()),    fuente));
                        PdfPCell c4 = new PdfPCell(new Phrase(safe(vac.getDescripcion()),        fuente));
                        for (PdfPCell c : new PdfPCell[] { c1, c2, c3, c4 }) {
                            c.setHorizontalAlignment(Element.ALIGN_CENTER);
                        }
                        tablaVacunas.addCell(c1);
                        tablaVacunas.addCell(c2);
                        tablaVacunas.addCell(c3);
                        tablaVacunas.addCell(c4);
                    }

                    document.add(tablaVacunas);
                }
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → que el controller decida 400 si corresponde
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar VacunacionUsuarios.pdf", e);
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

    // =========================
    // XLSX (nuevo)
    // =========================
    public byte[] generarReporteXLSX(ReporteVacunacionUsuariosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Vacunación");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges B1:Px (dos líneas de título)
            // 16 columnas totales -> columnas 0..15; B es 1 -> 1..15
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 15);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 15);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 15);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 15);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 15);

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? "REGISTRO DE VACUNACIÓN"
                    : request.getTitulo();
            UtilExcel.establecerTexto(hoja, 1, 1, UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "GÉNERO",
                    "CIUDAD", "SUCURSAL", "RÉGIMEN", "DEPARTAMENTO", "CARGO",
                    "ROL", "CORREO", "CARNET", "TIPO VACUNA", "FECHA", "DESCRIPCIÓN"
            };
            int[] anchos = {
                    10, 20, 20, 25, 15,
                    18, 18, 15, 20, 18,
                    14, 28, 12, 18, 15, 24
            };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupo → empleado → vacuna)
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            if (request.getDatos() != null) {
                for (AgrupadorVacunaUsuarioDTO grupo : request.getDatos()) {

                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoVacunaUsuarioDTO e : grupo.getEmpleados()) {

                        String apenom = (safe(e.getApellido()) + " " + safe(e.getNombre())).trim();

                        if (e.getVacunas() == null || e.getVacunas().isEmpty()) {
                            // fila "vacía" sin vacunas (si hace falta listar igual)
                            Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                            UtilExcel.establecerValor(r, 0, item++, null);
                            UtilExcel.establecerTexto(r, 1, safe(e.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, 2, safe(e.getCodigo()), null);
                            UtilExcel.establecerTexto(r, 3, apenom, null);
                            UtilExcel.establecerTexto(r, 4, safe(e.getGenero()), null);
                            UtilExcel.establecerTexto(r, 5, safe(e.getCiudad()), null);
                            UtilExcel.establecerTexto(r, 6, safe(e.getSucursal()), null);
                            UtilExcel.establecerTexto(r, 7, safe(e.getRegimen()), null);
                            UtilExcel.establecerTexto(r, 8, safe(e.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, 9, safe(e.getCargo()), null);
                            UtilExcel.establecerTexto(r, 10, safe(e.getRol()), null);
                            UtilExcel.establecerTexto(r, 11, safe(e.getCorreo()), null);
                            UtilExcel.establecerTexto(r, 12, "", null);
                            UtilExcel.establecerTexto(r, 13, "", null);
                            UtilExcel.establecerTexto(r, 14, "", null);
                            UtilExcel.establecerTexto(r, 15, "", null);
                            continue;
                        }

                        for (VacunaUsuarioDTO v : e.getVacunas()) {
                            Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                            UtilExcel.establecerValor(r, 0, item++, null);
                            UtilExcel.establecerTexto(r, 1, safe(e.getIdentificacion()), null);
                            UtilExcel.establecerTexto(r, 2, safe(e.getCodigo()), null);
                            UtilExcel.establecerTexto(r, 3, apenom, null);
                            UtilExcel.establecerTexto(r, 4, safe(e.getGenero()), null);
                            UtilExcel.establecerTexto(r, 5, safe(e.getCiudad()), null);
                            UtilExcel.establecerTexto(r, 6, safe(e.getSucursal()), null);
                            UtilExcel.establecerTexto(r, 7, safe(e.getRegimen()), null);
                            UtilExcel.establecerTexto(r, 8, safe(e.getDepartamento()), null);
                            UtilExcel.establecerTexto(r, 9, safe(e.getCargo()), null);
                            UtilExcel.establecerTexto(r, 10, safe(e.getRol()), null);
                            UtilExcel.establecerTexto(r, 11, safe(e.getCorreo()), null);

                            // carnet Sí/No (si el backend trae el campo; si no, queda vacío)
                            String carnetSN = (v.getCarnet() != null && !v.getCarnet().trim().isEmpty()) ? "Si" : "No";
                            UtilExcel.establecerTexto(r, 12, carnetSN, null);          // CARNET (Sí/No)
                            UtilExcel.establecerTexto(r, 13, safe(v.getTipo_vacuna()), null);
                            UtilExcel.establecerTexto(r, 14, fechaCortaExcel(v.getFecha()), null);
                            UtilExcel.establecerTexto(r, 15, safe(v.getDescripcion()), null);
                        }
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, encabezados.length - 1,
                        estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (zebra + AutoFilter)
            if (ultimaFila >= filaDatosInicio) {
                boolean[] filtros = new boolean[encabezados.length];
                for (int i = 0; i < filtros.length; i++)
                    filtros[i] = true;
                filtros[0] = false; // ITEM sin filtro

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "VacunasReporteTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
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

    private PdfPCell celdaSinBordeIzquierda(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setBorder(Rectangle.NO_BORDER);
        celda.setPaddingTop(6f);
        celda.setPaddingBottom(6f);
        return celda;
    }

    private String safe(String val) {
        return (val == null || val.equalsIgnoreCase("null")) ? "" : val;
    }

    private String formatearFecha(String fechaIso) {
        try {
            LocalDate fecha = LocalDate.parse(fechaIso.substring(0, 10));
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEE. dd/MM/yyyy", new Locale("es", "ES"));
            return formatter.format(fecha);
        } catch (Exception e) {
            return fechaIso;
        }
    }

    private String fechaCortaExcel(String iso) {
        if (iso == null)
            return "";
        String f = iso.trim();
        if (f.length() >= 10)
            return f.substring(0, 10);
        return f;
    }
}
