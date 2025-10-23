package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Empleado.EmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.Empleado.ReporteEmpleadosRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
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

    //METODO QUE GENERA EL REPORTE PDF
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
                request.getColorPrincipal()
            ));
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
            Color colorZebra     = ReporteUtil.colorZebraClaro();

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
                    tabla.addCell(ReporteUtil.crearCelda(e.getCodigo(),          ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getNombreCompleto(),  ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getIdentificacion(),  ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getFechaNacimiento(), ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getCorreo(),          ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getGenero(),          ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getEstadoCivil(),     ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getDomicilio(),       ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getTelefono(),        ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getEstadoTexto(),     ReporteUtil.fuenteTablaData(), bg));
                    tabla.addCell(ReporteUtil.crearCelda(e.getNacionalidad(),    ReporteUtil.fuenteTablaData(), bg));
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
    //          XLSX (LEGACY)
    // =========================
    public byte[] generarReporteEmpleadosXLSX(ReporteEmpleadosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet("Empleados");

            // 1) Logo A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            // 2) Merges B1:M1 ... B5:M5 (13 columnas totales A..M)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 12);
            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 12);
            UtilExcel.combinarCeldas(hoja, 2, 2, 1, 12);
            UtilExcel.combinarCeldas(hoja, 3, 3, 1, 12);
            UtilExcel.combinarCeldas(hoja, 4, 4, 1, 12);

            // 3) Títulos
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE EMPLEADOS", estiloTitulo);

            // 4) Encabezados + anchos
            final int filaEncabezado = 5; // A6
            String[] encabezados = {
                "ITEM","CODIGO","IDENTIFICACION","APELLIDO","NOMBRE","FECHA_NACIMIENTO",
                "ESTADO_CIVIL","GENERO","CORREO","ESTADO","DOMICILIO","TELEFONO","NACIONALIDAD"
            };
            int[] anchos = {10,20,20,20,20,20,20,20,20,20,20,20,20};

            Row rowHeader = UtilExcel.asegurarFila(hoja, filaEncabezado);
            for (int c = 0; c < encabezados.length; c++) {
                UtilExcel.establecerTexto(rowHeader, c, encabezados[c], null);
            }
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(rowHeader, encabezados.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, anchos);
            hoja.getRow(filaEncabezado).setHeightInPoints(18f);

            // 5) Cuerpo
            int filaDatosInicio = filaEncabezado + 1;
            int fila = filaDatosInicio;

            List<EmpleadoDTO> items = request.getEmpleados();
            if (items != null) {
                for (int i = 0; i < items.size(); i++) {
                    EmpleadoDTO e = items.get(i);

                    Row r = UtilExcel.asegurarFila(hoja, fila++);
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
                    UtilExcel.establecerValor(r,10, nvl(e.getDomicilio()), null);
                    UtilExcel.establecerValor(r,11, nvl(e.getTelefono()), null);
                    UtilExcel.establecerValor(r,12, nvl(e.getNacionalidad()), null);
                }
            }

            int ultimaFila = (fila == filaDatosInicio) ? filaEncabezado : (fila - 1);

            // 6) Alineación + bordes
            CellStyle centroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle izqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Header centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, filaEncabezado, filaEncabezado, 0, encabezados.length - 1, centroBorde, true);

            // Cuerpo: col 0 centrada; resto izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, centroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, 12, izqBorde, true);
            }

            // 7) Tabla estilizada + filtros (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                        hoja,
                        "Empleados",
                        filaEncabezado, 0,
                        ultimaFila, encabezados.length - 1,
                        true,
                        new boolean[]{false, true, true, true, true, true, true, true, true, true, true, true, true}
                );
            }

            libro.write(baos);
            return baos.toByteArray();
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    //          CSV (LEGACY)
    // =========================
    public byte[] generarReporteEmpleadosCSV(ReporteEmpleadosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            // Encabezados legacy (sin ITEM)
            sb.append("CODIGO,IDENTIFICACION,APELLIDO,NOMBRE,FECHA_NACIMIENTO,ESTADO_CIVIL,GENERO,CORREO,ESTADO,DOMICILIO,TELEFONO,NACIONALIDAD\n");

            List<EmpleadoDTO> items = request.getEmpleados();
            if (items != null) {
                for (EmpleadoDTO e : items) {

                    sb.append(csv(nvl(e.getCodigo()))).append(',')
                      .append(csv(nvl(e.getIdentificacion()))).append(',')
                      .append(csv((nvl(e.getApellido())))).append(',')
                      .append(csv((nvl(e.getNombre())))).append(',')
                      .append(csv(nvl(e.getFechaNacimiento()))).append(',')
                      .append(csv(nvl(e.getEstadoCivil()))).append(',')
                      .append(csv(nvl(e.getGenero()))).append(',')
                      .append(csv(nvl(e.getCorreo()))).append(',')
                      .append(csv(nvl(e.getEstadoTexto()))).append(',')
                      .append(csv(nvl(e.getDomicilio()))).append(',')
                      .append(csv(nvl(e.getTelefono()))).append(',')
                      .append(csv(nvl(e.getNacionalidad()))).append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    // =========================
    //           XML (LEGACY)
    // =========================
    public byte[] generarReporteEmpleadosXML(ReporteEmpleadosRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Empleados>\n");

            List<EmpleadoDTO> items = request.getEmpleados();
            if (items != null) {
                for (EmpleadoDTO e : items) {

                    sb.append("  <empleado codigo=\"").append(xml(nvl(e.getCodigo()))).append("\">\n");
                    sb.append("    <identificacion>").append(xml(nvl(e.getIdentificacion()))).append("</identificacion>\n");
                    sb.append("    <apellido>").append(xml(nvl(e.getApellido()))).append("</apellido>\n");
                    sb.append("    <nombre>").append(xml(nvl(e.getNombre()))).append("</nombre>\n");
                    sb.append("    <estadoCivil>").append(xml(nvl(e.getEstadoCivil()))).append("</estadoCivil>\n");
                    sb.append("    <genero>").append(xml(nvl(e.getGenero()))).append("</genero>\n");
                    sb.append("    <correo>").append(xml(nvl(e.getCorreo()))).append("</correo>\n");
                    sb.append("    <fechaNacimiento>").append(xml(nvl(e.getFechaNacimiento()))).append("</fechaNacimiento>\n");
                    sb.append("    <estado>").append(xml(nvl(e.getEstadoTexto()))).append("</estado>\n");
                    sb.append("    <domicilio>").append(xml(nvl(e.getDomicilio()))).append("</domicilio>\n");
                    sb.append("    <telefono>").append(xml(nvl(e.getTelefono()))).append("</telefono>\n");
                    sb.append("    <nacionalidad>").append(xml(nvl(e.getNacionalidad()))).append("</nacionalidad>\n");
                    sb.append("    <imagen>").append("") /* sin imagen en payload */ .append("</imagen>\n");
                    sb.append("  </empleado>\n");
                }
            }

            sb.append("</Empleados>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception ex) {
            ex.printStackTrace();
            return null;
        }
    }

    // =========================
    //              Helpers
    // =========================
    private static String nvl(String s) {
        return (s == null) ? "" : s;
    }


    private String csv(String v) {
        if (v == null) return "";
        boolean quote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return quote ? "\"" + s + "\"" : s;
    }

    private String xml(String v) {
        String s = (v == null) ? "" : v;
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"","&quot;")
                .replace("'","&apos;");
    }
}
