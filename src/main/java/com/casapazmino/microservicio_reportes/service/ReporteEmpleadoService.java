package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Empleado.EmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.Empleado.ReporteEmpleadosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.UtilXml;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.*;
import com.lowagie.text.pdf.*;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

@Service
public class ReporteEmpleadoService {

    // METODO QUE GENERA EL REPORTE PDF
    public byte[] generarReporteEmpleadosPDF(ReporteEmpleadosRequest request) {

        // DRY: constantes locales
        final float[] WIDTHS = { 1.8f, 5.5f, 3.7f, 3f, 7.3f, 2.7f, 3f, 3f, 3f, 2f, 3f };
        final String[] HEADERS = {
                "Código", "Nombre", "Identificación", "Fecha Nacimiento", "Correo",
                "Género", "Estado Civil", "Domicilio", "Teléfono", "Estado", "Nacionalidad"
        };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate());
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (helpers existentes)
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte("Lista de Empleados"));

            // Colores
            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorZebra = ReporteUtil.colorZebraClaro();

            // Tabla principal
            PdfPTable tabla = new PdfPTable(11);
            tabla.setWidthPercentage(100);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(10f);

            // Encabezados
            for (String h : HEADERS) {
                tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
            }

            // Cuerpo (zebra)
            List<EmpleadoDTO> empleados = request.getEmpleados();
            boolean zebra = false;
            if (empleados != null) {
                for (EmpleadoDTO e : empleados) {
                    Color bg = zebra ? colorZebra : Color.WHITE;
                    tabla.addCell(ReporteUtil.crearCelda(e.getCodigo(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getNombreCompleto(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getIdentificacion(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getFechaNacimiento(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getCorreo(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getGenero(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getEstadoCivil(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getDomicilio(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getTelefono(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getEstadoTexto(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getNacionalidad(), ReporteUtil.fuenteTablaData(), bg));
                    zebra = !zebra;
                }
            }

            document.add(tabla);

            // 3) Cierre y retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // si un helper valida y falla, que el controller decida (posible 400)
            throw e;
        } catch (Exception e) {
            // 500 interno uniforme
            throw new ReportBuildException("No se pudo generar ReporteEmpleados.pdf", e);
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
    // XLSX
    // =========================
    public byte[] generarReporteEmpleadosXLSX(ReporteEmpleadosRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA = "Empleados";
        final int FILA_ENCABEZADO = 5;

        // Merges B1:M1 ... B5:M5 => (row 0..4, col 1..12)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 12;

        final String[] HEADERS = {
                "ITEM", "CODIGO", "IDENTIFICACION", "APELLIDO", "NOMBRE", "FECHA_NACIMIENTO",
                "ESTADO_CIVIL", "GENERO", "CORREO", "ESTADO", "DOMICILIO", "TELEFONO", "NACIONALIDAD"
        };
        final int[] ANCHOS = { 10, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20 };

        try (XSSFWorkbook libro = new XSSFWorkbook();
                ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);
            }

            // 2) Merges B1:M1 ... B5:M5
            for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE EMPLEADOS", estiloTitulo);

            // 4) Encabezados + anchos
            Row rowHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(rowHeader, c, HEADERS[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(rowHeader, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) Cuerpo
            int filaDatosIni = FILA_ENCABEZADO + 1;
            int filaAct = filaDatosIni;

            List<EmpleadoDTO> items = request.getEmpleados();
            if (items != null) {
                for (int i = 0; i < items.size(); i++) {
                    EmpleadoDTO e = items.get(i);
                    Row r = UtilExcel.asegurarFila(hoja, filaAct++);
                    UtilExcel.establecerValor(r, 0, i + 1, null); // ITEM
                    UtilExcel.establecerValor(r, 1, nvl(e.getCodigo()), null);
                    UtilExcel.establecerValor(r, 2, nvl(e.getIdentificacion()), null);
                    UtilExcel.establecerValor(r, 3, nvl(e.getApellido()), null);
                    UtilExcel.establecerValor(r, 4, nvl(e.getNombre()), null);
                    UtilExcel.establecerValor(r, 5, nvl(e.getFechaNacimiento()), null);
                    UtilExcel.establecerValor(r, 6, nvl(e.getEstadoCivil()), null);
                    UtilExcel.establecerValor(r, 7, nvl(e.getGenero()), null);
                    UtilExcel.establecerValor(r, 8, nvl(e.getCorreo()), null);
                    UtilExcel.establecerValor(r, 9, nvl(e.getEstadoTexto()), null);
                    UtilExcel.establecerValor(r, 10, nvl(e.getDomicilio()), null);
                    UtilExcel.establecerValor(r, 11, nvl(e.getTelefono()), null);
                    UtilExcel.establecerValor(r, 12, nvl(e.getNacionalidad()), null);
                }
            }

            int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

            // 6) Alineación + bordes
            CellStyle centroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle izqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, centroBorde,
                    true);

            // Cuerpo: col 0 centrada; resto izquierda
            if (ultimaFila >= filaDatosIni) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, centroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, 12, izqBorde, true);
            }

            // 7) Tabla estilizada + filtros (ITEM sin filtro)
            if (ultimaFila >= filaDatosIni) {
                boolean[] filtros = new boolean[] {
                        false, true, true, true, true, true, true, true, true, true, true, true, true
                };
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "Empleados",
                        FILA_ENCABEZADO, 0,
                        ultimaFila, HEADERS.length - 1,
                        true,
                        filtros);
            }

            // 8) Cierre + retorno
            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e; // Validación → 400
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar Empleados.xlsx", e); // Interno → 500
        }
    }

    // =========================
    // CSV (LEGACY)
    // =========================
    public byte[] generarReporteEmpleadosCSV(ReporteEmpleadosRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Empleados.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = {
                "CODIGO", "IDENTIFICACION", "APELLIDO", "NOMBRE", "FECHA_NACIMIENTO", "ESTADO_CIVIL",
                "GENERO", "CORREO", "ESTADO", "DOMICILIO", "TELEFONO", "NACIONALIDAD"
        };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0)
                    sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            List<EmpleadoDTO> items = request.getEmpleados();
            if (items != null && !items.isEmpty()) {
                for (EmpleadoDTO e : items) {
                    String codigo = (e.getCodigo() == null) ? "" : e.getCodigo();
                    String identificacion = (e.getIdentificacion() == null) ? "" : e.getIdentificacion();
                    String apellido = (e.getApellido() == null) ? "" : e.getApellido();
                    String nombre = (e.getNombre() == null) ? "" : e.getNombre();
                    String fechaNac = (e.getFechaNacimiento() == null) ? "" : e.getFechaNacimiento();
                    String estadoCivil = (e.getEstadoCivil() == null) ? "" : e.getEstadoCivil();
                    String genero = (e.getGenero() == null) ? "" : e.getGenero();
                    String correo = (e.getCorreo() == null) ? "" : e.getCorreo();
                    String estadoTexto = (e.getEstadoTexto() == null) ? "" : e.getEstadoTexto();
                    String domicilio = (e.getDomicilio() == null) ? "" : e.getDomicilio();
                    String telefono = (e.getTelefono() == null) ? "" : e.getTelefono();
                    String nacionalidad = (e.getNacionalidad() == null) ? "" : e.getNacionalidad();

                    sb.append(UtilCsv.csvEscape(codigo)).append(DELIM)
                            .append(UtilCsv.csvEscape(identificacion)).append(DELIM)
                            .append(UtilCsv.csvEscape(apellido)).append(DELIM)
                            .append(UtilCsv.csvEscape(nombre)).append(DELIM)
                            .append(UtilCsv.csvEscape(fechaNac)).append(DELIM)
                            .append(UtilCsv.csvEscape(estadoCivil)).append(DELIM)
                            .append(UtilCsv.csvEscape(genero)).append(DELIM)
                            .append(UtilCsv.csvEscape(correo)).append(DELIM)
                            .append(UtilCsv.csvEscape(estadoTexto)).append(DELIM)
                            .append(UtilCsv.csvEscape(domicilio)).append(DELIM)
                            .append(UtilCsv.csvEscape(telefono)).append(DELIM)
                            .append(UtilCsv.csvEscape(nacionalidad)).append(EOL);
                }
            }

            // Retorno (nunca null)
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Interno → 500
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    // =========================
    // XML (LEGACY)
    // =========================
    public byte[] generarReporteEmpleadosXML(ReporteEmpleadosRequest request) {
        final String NOMBRE_REPORTE = "Empleados.xml";
        final String ROOT_TAG = "Empleados";
        final String ITEM_TAG = "empleado";
        final String EOL = "\n";
        final String IND = "  ";

        try {
            StringBuilder sb = new StringBuilder(8_192);

            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>").append(EOL);
            sb.append("<").append(ROOT_TAG).append(">").append(EOL);

            List<EmpleadoDTO> items = request.getEmpleados();
            if (items == null || items.isEmpty()) {
                sb.append(IND).append("<lista>NO DEFINIDO</lista>").append(EOL);
            } else {
                for (EmpleadoDTO e : items) {
                    sb.append(IND).append("<").append(ITEM_TAG)
                            .append(" codigo=\"").append(UtilXml.xmlEsc(e.getCodigo())).append("\">").append(EOL);

                    sb.append(IND).append(IND).append("<identificacion>")
                            .append(UtilXml.xmlEsc(e.getIdentificacion()))
                            .append("</identificacion>").append(EOL);

                    sb.append(IND).append(IND).append("<apellido>")
                            .append(UtilXml.xmlEsc(e.getApellido()))
                            .append("</apellido>").append(EOL);

                    sb.append(IND).append(IND).append("<nombre>")
                            .append(UtilXml.xmlEsc(e.getNombre()))
                            .append("</nombre>").append(EOL);

                    sb.append(IND).append(IND).append("<estadoCivil>")
                            .append(UtilXml.xmlEsc(e.getEstadoCivil()))
                            .append("</estadoCivil>").append(EOL);

                    sb.append(IND).append(IND).append("<genero>")
                            .append(UtilXml.xmlEsc(e.getGenero()))
                            .append("</genero>").append(EOL);

                    sb.append(IND).append(IND).append("<correo>")
                            .append(UtilXml.xmlEsc(e.getCorreo()))
                            .append("</correo>").append(EOL);

                    sb.append(IND).append(IND).append("<fechaNacimiento>")
                            .append(UtilXml.xmlEsc(e.getFechaNacimiento()))
                            .append("</fechaNacimiento>").append(EOL);

                    sb.append(IND).append(IND).append("<estado>")
                            .append(UtilXml.xmlEsc(e.getEstadoTexto()))
                            .append("</estado>").append(EOL);

                    sb.append(IND).append(IND).append("<domicilio>")
                            .append(UtilXml.xmlEsc(e.getDomicilio()))
                            .append("</domicilio>").append(EOL);

                    sb.append(IND).append(IND).append("<telefono>")
                            .append(UtilXml.xmlEsc(e.getTelefono()))
                            .append("</telefono>").append(EOL);

                    sb.append(IND).append(IND).append("<nacionalidad>")
                            .append(UtilXml.xmlEsc(e.getNacionalidad()))
                            .append("</nacionalidad>").append(EOL);

                    sb.append(IND).append(IND).append("<imagen>")
                            .append("") // sin imagen en payload (se deja vacío)
                            .append("</imagen>").append(EOL);

                    sb.append(IND).append("</").append(ITEM_TAG).append(">").append(EOL);
                }
            }

            sb.append("</").append(ROOT_TAG).append(">").append(EOL);
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

    private static String nvl(String s) {
        return (s == null) ? "" : s;
    }

}
