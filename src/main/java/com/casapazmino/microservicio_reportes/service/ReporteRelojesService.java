package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Dispositivo.RelojDTO;
import com.casapazmino.microservicio_reportes.model.Dispositivo.ReporteRelojesRequest;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

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
public class ReporteRelojesService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReportePDF(ReporteRelojesRequest request) {

        // ➊ DRY: constantes locales (no cambia look & feel)
        final String TITULO = "LISTA DE DISPOSITIVOS";
        final float[] WIDTHS = {
            2.5f, 4f, 2.2f, 4f, 3.7f, 2.5f, 3f, 2f, 2f, 2.5f, 3f, 1.5f, 3.5f, 3f, 4.5f
        };
        final String[] HEADERS = {
            "Código", "Empresa", "Ciudad", "Establecimiento", "Departamento", "Nombre", "IP",
            "Puerto", "Marca", "Modelo", "Serie", "Mac", "ID Fabricación", "Fabricante", "Zona Horaria"
        };
        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_BEFORE = 10f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_ZEBRA    = ReporteUtil.colorZebraClaro();

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
            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            // Títulos
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO));

            // Tabla principal
            PdfPTable tabla = new PdfPTable(HEADERS.length);
            tabla.setWidthPercentage(WIDTH_PERCENT_100);
            tabla.setWidths(WIDTHS);
            tabla.setSpacingBefore(SPACING_BEFORE);

            // Encabezados
            for (String h : HEADERS) {
                tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));
            }

            // Cuerpo con zebra
            boolean zebra = false;
            List<RelojDTO> lista = request.getRelojes();
            if (lista != null) {
                for (RelojDTO r : lista) {
                    Color fondo = zebra ? COLOR_ZEBRA : Color.WHITE;
                    zebra = !zebra;

                    tabla.addCell(ReporteUtil.crearCelda(r.getCodigo(),                     fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomempresa(),                 fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomciudad(),                  fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomsucursal(),                fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNomdepar(),                   fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getNombre(),                     fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getIp(),                         fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(String.valueOf(r.getPuerto()),     fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getMarca(),                      fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getModelo(),                     fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getSerie(),                      fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getMac(),                        fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getIdFabricacion(),              fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(r.getFabricante(),                 fuenteAuxiliarTabla(), fondo));
                    tabla.addCell(ReporteUtil.crearCelda(
                            r.getZonaHorariaDispositivo() + " (" + r.getFormatoGmtDispositivo() + ")",
                            fuenteAuxiliarTabla(), fondo));
                }
            }

            document.add(tabla);

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → el controller decidirá 400 si aplica
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteRelojes.pdf", e);
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

    public Font fuenteAuxiliarTabla() {
        return FontFactory.getFont(FontFactory.HELVETICA, 7);
    }

    private String nv(Object v) {
        return v == null ? "" : String.valueOf(v);
    }

    // =========================
    // XLSX 
    // =========================
    public byte[] generarReporteXLSX(ReporteRelojesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA    = "Relojes"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;        // fila visual 6 (idx 5)

        // MERGES exactos (B1:U1 ... B5:U5) => (row 0..4, col 1..20)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 20;

        final String[] HEADERS = {
            "ITEM", "ID", "CODIGO", "NOMBRE", "IP", "PUERTO", "CONTRASEÑA",
            "MARCA", "MODELO", "SERIE", "ID_FABRICACION", "FABRICANTE", "MAC",
            "TIPO CONEXION", "ID SUCURSAL", "ID DEPARTAMENTO",
            "NOMBRE DEPARTAMENTO", "NOMBRE CIUDAD", "TEMPERATURA",
            "ZONA HORARIA DISPOSITIVO", "FORMATO GMT DISPOSITIVO"
        };

        final int[] ANCHOS = {
            10, 20, 20, 20, 20, 15, 20, 20, 20, 20,
            20, 20, 20, 20, 20, 20, 25, 25, 15, 30, 30
        };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] {
            false, true, true, true, true, true, true,
            true,  true, true, true, true, true, true,
            true,  true, true, true, true, true, true
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

            // 2) MERGES exactos (B1:U1 ... B5:U5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                hoja, 0, 1,
                UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                estiloTitulo
            ); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "LISTA DE RELOJES", estiloTitulo); // B2

            // 4) ENCABEZADOS + ANCHOS (fila 6 → idx 5)
            Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int c = 0; c < HEADERS.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
            }
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloEncabezado);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

            // 5) CUERPO
            int filaDatosInicio = FILA_ENCABEZADO + 1;
            int filaActual = filaDatosInicio;

            List<RelojDTO> relojes = request.getRelojes();
            if (relojes != null) {
                for (int i = 0; i < relojes.size(); i++) {
                    RelojDTO r = relojes.get(i);
                    Row row = UtilExcel.asegurarFila(hoja, filaActual++);

                    UtilExcel.establecerValor(row, 0,  i + 1,           null); // ITEM
                    UtilExcel.establecerValor(row, 1,  r.getId(),       null);
                    UtilExcel.establecerValor(row, 2,  nv(r.getCodigo()),             null);
                    UtilExcel.establecerValor(row, 3,  nv(r.getNombre()),             null);
                    UtilExcel.establecerValor(row, 4,  nv(r.getIp()),                 null);
                    UtilExcel.establecerValor(row, 5,  r.getPuerto(),                 null);
                    UtilExcel.establecerValor(row, 6,  nv(r.getContrasenia()),        null);
                    UtilExcel.establecerValor(row, 7,  nv(r.getMarca()),              null);
                    UtilExcel.establecerValor(row, 8,  nv(r.getModelo()),             null);
                    UtilExcel.establecerValor(row, 9,  nv(r.getSerie()),              null);
                    UtilExcel.establecerValor(row, 10, nv(r.getIdFabricacion()),      null);
                    UtilExcel.establecerValor(row, 11, nv(r.getFabricante()),         null);
                    UtilExcel.establecerValor(row, 12, nv(r.getMac()),                null);
                    UtilExcel.establecerValor(row, 13, nv(r.getTipoConexion()),       null);
                    UtilExcel.establecerValor(row, 14, r.getIdSucursal(),             null);
                    UtilExcel.establecerValor(row, 15, r.getIdDepartamento(),         null);
                    UtilExcel.establecerValor(row, 16, nv(r.getNomdepar()),           null);
                    UtilExcel.establecerValor(row, 17, nv(r.getNomciudad()),          null);
                    UtilExcel.establecerValor(row, 18, nv(r.getTemperatura()),        null);
                    UtilExcel.establecerValor(row, 19, nv(r.getZonaHorariaDispositivo()), null);
                    UtilExcel.establecerValor(row, 20, nv(r.getFormatoGmtDispositivo()),  null);
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) ALINEACIONES + BORDES (header centrado; cuerpo col 0 centrado, resto izquierda)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(
                hoja, FILA_ENCABEZADO, FILA_ENCABEZADO,
                0, HEADERS.length - 1, estiloCentroBorde, true
            );

            // Cuerpo: col 0 centrada; col 1..20 izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0,  estiloCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde, true);
            }

            // 7) TABLA estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "RelojesTabla",
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
            // Validación → 400 (lo maneja el controller)
            throw e;
        } catch (Exception e) {
            // Interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Relojes.xlsx", e);
        }
    }

    
    // =========================
    // CSV (orden estable, similar a Excel)
    // =========================
    public byte[] generarReporteCSV(ReporteRelojesRequest request) {
        try {
            String[] headers = {
                    "id", "codigo", "nombre", "ip", "puerto", "contrasenia", "marca", "modelo", "serie",
                    "id_fabricacion", "fabricante", "mac", "tipo_conexion", "id_sucursal", "id_departamento",
                    "nomdepar", "nomciudad", "temperatura", "zona_horaria_dispositivo", "formato_gmt_dispositivo",
                    "nomempresa", "nomsucursal"
            };

            StringBuilder sb = new StringBuilder();
            // encabezados
            for (int i = 0; i < headers.length; i++) {
                sb.append(headers[i]);
                if (i < headers.length - 1)
                    sb.append(',');
            }
            sb.append('\n');

            List<RelojDTO> items = request.getRelojes();
            if (items != null) {
                for (RelojDTO r : items) {
                    String[] vals = {
                            sv(r.getId()),
                            nv(r.getCodigo()),
                            nv(r.getNombre()),
                            nv(r.getIp()),
                            sv(r.getPuerto()),
                            nv(r.getContrasenia()),
                            nv(r.getMarca()),
                            nv(r.getModelo()),
                            nv(r.getSerie()),
                            nv(r.getIdFabricacion()),
                            nv(r.getFabricante()),
                            nv(r.getMac()),
                            nv(r.getTipoConexion()),
                            sv(r.getIdSucursal()),
                            sv(r.getIdDepartamento()),
                            nv(r.getNomdepar()),
                            nv(r.getNomciudad()),
                            nv(r.getTemperatura()),
                            nv(r.getZonaHorariaDispositivo()),
                            nv(r.getFormatoGmtDispositivo()),
                            nv(r.getNomempresa()),
                            nv(r.getNomsucursal())
                    };
                    for (int i = 0; i < vals.length; i++) {
                        sb.append(csv(vals[i]));
                        if (i < vals.length - 1)
                            sb.append(',');
                    }
                    sb.append('\n');
                }
            }
            return sb.toString().getBytes(StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // =========================
    // XML (igual a tu legacy front)
    // =========================
    public byte[] generarReporteXML(ReporteRelojesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Relojes>\n");

            List<RelojDTO> items = request.getRelojes();
            if (items != null) {
                for (RelojDTO r : items) {
                    sb.append("  <reloj id=\"").append(xml(r.getId())).append("\">\n");
                    sb.append("    <codigo>").append(xml(r.getCodigo())).append("</codigo>\n");
                    sb.append("    <nombre_empresa>").append(xml(r.getNomempresa())).append("</nombre_empresa>\n");
                    sb.append("    <nombre_ciudad>").append(xml(r.getNomciudad())).append("</nombre_ciudad>\n");
                    sb.append("    <nombre_sucursal>").append(xml(r.getNomsucursal())).append("</nombre_sucursal>\n");
                    sb.append("    <nombre_departamento>").append(xml(r.getNomdepar()))
                            .append("</nombre_departamento>\n");
                    sb.append("    <nombre>").append(xml(r.getNombre())).append("</nombre>\n");
                    sb.append("    <ip>").append(xml(r.getIp())).append("</ip>\n");
                    sb.append("    <puerto>").append(xml(r.getPuerto())).append("</puerto>\n");
                    sb.append("    <marca>").append(xml(r.getMarca())).append("</marca>\n");
                    sb.append("    <modelo>").append(xml(r.getModelo())).append("</modelo>\n");
                    sb.append("    <serie>").append(xml(r.getSerie())).append("</serie>\n");
                    sb.append("    <mac>").append(xml(r.getMac())).append("</mac>\n");
                    sb.append("    <id_fabricacion>").append(xml(r.getIdFabricacion())).append("</id_fabricacion>\n");
                    sb.append("    <fabricante>").append(xml(r.getFabricante())).append("</fabricante>\n");
                    sb.append("    <zona_horaria>")
                            .append(xml(
                                    (nv(r.getZonaHorariaDispositivo()) + " (" + nv(r.getFormatoGmtDispositivo()) + ")")
                                            .trim()))
                            .append("</zona_horaria>\n");
                    sb.append("  </reloj>\n");
                }
            }

            sb.append("</Relojes>\n");
            return sb.toString().getBytes(StandardCharsets.UTF_8);

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // ===== Helpers CSV / XML =====
    private String sv(Number n) {
        return n == null ? "" : String.valueOf(n);
    }

    private String csv(String v) {
        if (v == null)
            return "";
        boolean quote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return quote ? "\"" + s + "\"" : s;
    }

    private String xml(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

}
