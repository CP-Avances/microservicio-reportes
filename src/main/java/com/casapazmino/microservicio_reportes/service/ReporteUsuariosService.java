package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteUsuario.AgrupadorUsuariosDTO;
import com.casapazmino.microservicio_reportes.model.ReporteUsuario.ReporteUsuariosRequest;
import com.casapazmino.microservicio_reportes.model.ReporteUsuario.UsuarioDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
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
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            Document document = new Document(PageSize.A4.rotate(), 40, 40, 60, 40);
            PdfWriter writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()
            ));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            Paragraph empresa = new Paragraph(request.getEmpresa(), ReporteUtil.fuenteEncabezado());
            empresa.setAlignment(Element.ALIGN_CENTER);
            empresa.setSpacingAfter(5f);
            document.add(empresa);

            Paragraph titulo = new Paragraph(request.getTitulo(), ReporteUtil.fuenteEncabezado());
            titulo.setAlignment(Element.ALIGN_CENTER);
            titulo.setSpacingAfter(10f);
            document.add(titulo);

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Font fuente = ReporteUtil.fuenteTexto();

            for (AgrupadorUsuariosDTO grupo : request.getDatos()) {
                String descripcion = "";
                String establecimiento = safe("SUCURSAL: " + grupo.getSucursal());

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

                PdfPTable cabecera = new PdfPTable(3);
                cabecera.setWidthPercentage(100);
                cabecera.setWidths(new float[]{3, 3, 2});
                cabecera.setSpacingBefore(10f);
                cabecera.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                cabecera.addCell(celdaSinBorde(descripcion, fuente, colorSecundario));
                cabecera.addCell(celdaSinBorde(establecimiento, fuente, colorSecundario));
                cabecera.addCell(celdaSinBorde("N° Registros: " + grupo.getEmpleados().size(), fuente, colorSecundario));

                cabecera.setTableEvent((table, widths, heights, headerRows, rowStart, canvas) -> {
                    PdfContentByte cb = canvas[PdfPTable.LINECANVAS];
                    cb.rectangle(
                            widths[0][0],
                            heights[heights.length - 1],
                            widths[0][widths[0].length - 1] - widths[0][0],
                            heights[0] - heights[heights.length - 1]
                    );
                    cb.stroke();
                });

                document.add(cabecera);

                // Tabla de usuarios
                PdfPTable tablaUsuarios = new PdfPTable(14);
                tablaUsuarios.setWidthPercentage(100);
                tablaUsuarios.setSpacingBefore(0f);
                tablaUsuarios.setWidths(new float[]{1, 3, 3, 4, 3, 2, 3, 3, 3, 3, 3, 3, 3, 5});

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

                int index = 1;
                for (UsuarioDTO usu : grupo.getEmpleados()) {
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(String.valueOf(index++), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getIdentificacion()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaCentro(safe(usu.getCodigo()), fuente));
                    tablaUsuarios.addCell(ReporteUtil.celdaIzquierda(safe(usu.getApellido()) + " " + safe(usu.getNombre()), fuente));
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
    public byte[] generarReporteXLSX(ReporteUsuariosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Usuarios");

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5

            // 2) Merges B1:Lx (replicando ExcelJS)
            // B = col 1; L = col 11 (0-based). Filas 0..4.
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 11);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 11);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 11);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 11);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 11);

            // 3) Títulos centrados (Empresa + Título del reporte)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            String titulo = (request.getTitulo() == null || request.getTitulo().isEmpty())
                    ? "LISTA DE USUARIOS"
                    : request.getTitulo();
            UtilExcel.establecerTexto(hoja, 1, 1, UtilExcel.aMayusculasSeguras(titulo), estiloTitulo);

            // 4) Encabezados + anchos (fila 6 → idx 5)
            final int filaEncabezado = 5;
            String[] encabezados = {
                    "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO", "NOMBRE",
                    "USUARIO", "GÉNERO", "NACIONALIDAD", "CIUDAD", "SUCURSAL",
                    "RÉGIMEN", "DEPARTAMENTO", "CARGO", "ROL", "CORREO"
            };
            int[] anchos = {
                    10, 20, 20, 30, 20,
                    20, 20, 20, 20, 20,
                    20, 20, 20, 30, 35
            };

            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, encabezados[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, encabezados.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo (aplanado grupo → empleados)
            int filaDatosInicio = filaEncabezado + 1;
            int filaActual = filaDatosInicio;
            int item = 1;

            if (request.getDatos() != null) {
                for (AgrupadorUsuariosDTO grupo : request.getDatos()) {
                    if (grupo.getEmpleados() == null) continue;

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

            int ultimaFila = (filaActual == filaDatosInicio) ? filaEncabezado : (filaActual - 1);

            // 6) Alineaciones + bordes
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde   = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1,
                    estiloCentroBorde, true);

            if (ultimaFila >= filaDatosInicio) {
                // ITEM centrado
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estiloCentroBorde, true);
                // resto izquierda (como en tu bucle de ExcelJS)
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, encabezados.length - 1,
                        estiloIzqBorde, true);
            }

            // 7) Tabla estilizada (zebra + AutoFilter)
            if (ultimaFila >= filaDatosInicio) {
                boolean[] filtros = new boolean[encabezados.length];
                for (int i = 0; i < filtros.length; i++) filtros[i] = true;
                filtros[0] = false; // ITEM sin filtro

                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "UsuariosexcelTabla",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        filtros
                );
            }

            libro.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
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
