package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteUsuario.AgrupadorUsuariosDTO;
import com.casapazmino.microservicio_reportes.model.ReporteUsuario.ReporteUsuariosRequest;
import com.casapazmino.microservicio_reportes.model.ReporteUsuario.UsuarioDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Service
public class ReporteUsuariosService {

    public byte[] generarReportePDF(ReporteUsuariosRequest request) {
        // DRY: constantes locales
        final float[] WIDTHS_CABECERA = { 3f, 3f, 2f };
        final float[] WIDTHS_TABLA_USU = { 1f, 4f, 2.5f, 4f, 3f, 3f, 4f, 3f, 3f, 3f, 5f, 3f, 3f, 5f };
        final String PREF_SUCURSAL = "SUCURSAL: ";
        final String PREF_REGISTROS = "N° Registros: ";
        final float MARGEN_IZQ = 40f, MARGEN_DER = 40f, MARGEN_SUP = 60f, MARGEN_INF = 40f;

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), MARGEN_IZQ, MARGEN_DER, MARGEN_SUP, MARGEN_INF);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (respetando tu diseño/estilos actuales)
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Encabezados empresa/título (manteniendo Paragraph + fuente existente)
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(safe(request.getTitulo())));


            // Colores y fuente (calculados una vez)
            final Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            final Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            final Font fuente = ReporteUtil.fuenteTexto();

            // Por cada grupo (según filtro)
            for (AgrupadorUsuariosDTO grupo : request.getDatos()) {
                // Descripción de bloque según tipo de filtro
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

                // Cabecera del bloque (sin bordes en celdas, con borde externo via TableEvent)
                PdfPTable cabecera = new PdfPTable(3);
                cabecera.setWidthPercentage(100);
                cabecera.setWidths(WIDTHS_CABECERA);
                cabecera.setSpacingBefore(10f);
                cabecera.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                cabecera.addCell(celdaSinBorde(descripcion, fuente, colorSecundario));
                cabecera.addCell(celdaSinBorde(establecimiento, fuente, colorSecundario));
                cabecera.addCell(celdaSinBorde(PREF_REGISTROS + grupo.getEmpleados().size(), fuente, colorSecundario));

                // Borde externo del bloque (respetando tu implementación)
                cabecera.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                    PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                    cb.rectangle(
                            widths[0][0],
                            heights[heights.length - 1],
                            widths[0][widths[0].length - 1] - widths[0][0],
                            heights[0] - heights[heights.length - 1]);
                    cb.stroke();
                });

                document.add(cabecera);

                // Tabla de usuarios
                PdfPTable tablaUsuarios = new PdfPTable(14);
                tablaUsuarios.setWidthPercentage(100);
                tablaUsuarios.setSpacingBefore(0f);
                tablaUsuarios.setWidths(WIDTHS_TABLA_USU);

                // Encabezados
                tablaUsuarios.addCell(ReporteUtil.crearCelda("N°", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("IDENTIFICACIÓN", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CÓDIGO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("EMPLEADO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("USUARIO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("GÉNERO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("NACIONALIDAD", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CIUDAD", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("SUCURSAL", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("RÉGIMEN", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("DEPARTAMENTO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CARGO", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("ROL", fuente, colorPrincipal));
                tablaUsuarios.addCell(ReporteUtil.crearCelda("CORREO", fuente, colorPrincipal));

                // Cuerpo
                int index = 1;
                for (UsuarioDTO usu : grupo.getEmpleados()) {
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(String.valueOf(index++), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getIdentificacion()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCodigo()), fuente));
                    tablaUsuarios.addCell(
                            ReporteUtil.celdaCentro(safe(usu.getApellido()) + " " + safe(usu.getNombre()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getUsuario()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getGenero()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getNacionalidad()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCiudad()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getSucursal()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getRegimen()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getDepartamento()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCargo()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getRol()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaIzquierda(safe(usu.getCorreo()), fuente));
                }

                document.add(tablaUsuarios);
            }

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Si algún helper valida y lanza IAEx → que el controller decida 400
            throw e;
        } catch (Exception e) {
            // Fallo interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Usuarios.pdf", e);
        } finally {
            // 4) Ciclo de recursos garantizado
            if (document != null && document.isOpen()) {
                try {
                    document.close();
                } catch (Exception ignore) {
                }
            }
            if (writer != null) {
                try {
                    writer.close();
                } catch (Exception ignore) {
                }
            }
            if (baos != null) {
                try {
                    baos.close();
                } catch (Exception ignore) {
                }
            }
        }
    }

    // =========================
    // XLSX (nuevo)
    // =========================
    public byte[] generarReporteXLSX(ReporteUsuariosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Usuarios"; // ≤ 31 chars
        final int FILA_ENCABEZADO = 5; // fila 6 (idx 5)

        // MERGES exactos (B1:L1 ... B5:L5) => (row 0..4, col 1..11)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 11;

        final String[] HEADERS = {
                "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO", "NOMBRE",
                "USUARIO", "GÉNERO", "NACIONALIDAD", "CIUDAD", "SUCURSAL",
                "RÉGIMEN", "DEPARTAMENTO", "CARGO", "ROL", "CORREO"
        };

        final int[] ANCHOS = {
                10, 20, 20, 30, 20,
                20, 20, 20, 20, 20,
                20, 20, 20, 30, 35
        };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] {
                false, true, true, true, true,
                true, true, true, true, true,
                true, true, true, true, true
        };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:L1 ... B5:L5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos centrados (Empresa + Título)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                    hoja, 0, 1,
                    UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                    estiloTitulo); // B1
            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? "LISTA DE USUARIOS"
                    : request.getTitulo();
            UtilExcel.establecerTexto(hoja, 1, 1, UtilExcel.aMayusculasSeguras(titulo), estiloTitulo); // B2

            // 4) Encabezados + anchos (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupo → empleados)
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            if (request.getDatos() != null) {
                for (AgrupadorUsuariosDTO grupo : request.getDatos()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (UsuarioDTO u : grupo.getEmpleados()) {
                        Row r = UtilExcel.asegurarFila(hoja, filaActual++);
                        UtilExcel.establecerValor(r, 0, item++, null);
                        UtilExcel.establecerTexto(r, 1, safe(u.getIdentificacion()), null);
                        UtilExcel.establecerTexto(r, 2, safe(u.getCodigo()), null);
                        UtilExcel.establecerTexto(r, 3, safe(u.getApellido()), null);
                        UtilExcel.establecerTexto(r, 4, safe(u.getNombre()), null);
                        UtilExcel.establecerTexto(r, 5, safe(u.getUsuario()), null);
                        UtilExcel.establecerTexto(r, 6, safe(u.getGenero()), null);
                        UtilExcel.establecerTexto(r, 7, safe(u.getNacionalidad()), null);
                        UtilExcel.establecerTexto(r, 8, safe(u.getCiudad()), null);
                        UtilExcel.establecerTexto(r, 9, safe(u.getSucursal()), null);
                        UtilExcel.establecerTexto(r, 10, safe(u.getRegimen()), null);
                        UtilExcel.establecerTexto(r, 11, safe(u.getDepartamento()), null);
                        UtilExcel.establecerTexto(r, 12, safe(u.getCargo()), null);
                        UtilExcel.establecerTexto(r, 13, safe(u.getRol()), null);
                        UtilExcel.establecerTexto(r, 14, safe(u.getCorreo()), null);
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) Alineaciones + bordes (header centrado; cuerpo col 0 centrada, resto
            // izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado
            UtilExcel.aplicarEstiloARegion(
                    hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                    0, HEADERS.length - 1, estiloCentroBorde, true);

            // Cuerpo
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true); // ITEM
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde,
                        true);
            }

            // 7) Tabla estilizada (zebra + AutoFilter) — ITEM sin filtro
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "UsuariosexcelTabla",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        FILTROS);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones → 400
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar Usuarios.xlsx", e);
        }
    }

    private PdfPCell celdaSinBorde(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuente));
        celda.setBackgroundColor(fondo);
        celda.setBorder(Rectangle.NO_BORDER);
        celda.setPaddingTop(6f);
        celda.setPaddingBottom(6f);
        return celda;
    }

    private String safe(String val) {
        return (val == null || val.equalsIgnoreCase("null")) ? "" : val;
    }
}
