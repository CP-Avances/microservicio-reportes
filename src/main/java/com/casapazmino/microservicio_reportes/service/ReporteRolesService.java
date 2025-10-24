package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Rol.FuncionDTO;
import com.casapazmino.microservicio_reportes.model.Rol.ReporteRolesRequest;
import com.casapazmino.microservicio_reportes.model.Rol.RolDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilCsv;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;

import com.lowagie.text.Document;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Phrase;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Row; // <-- ESTE
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Service
public class ReporteRolesService {

    // METODO QUE GENERA EL PDF
    public byte[] generarReporteRolesPDF(ReporteRolesRequest request) {

        // ➊ DRY: constantes locales
        final String TITULO = "PERMISOS O FUNCIONALIDADES DEL ROL";
        final float[] WIDTHS_INFO = { 3f, 4f, 4f, 2f, 2f };
        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_BEFORE_ENCABEZADO = 10f;
        final float SPACING_AFTER_ENCABEZADO = 5f;
        final float SPACING_BEFORE_TABLA = 5f;
        final float PADDING_TITULOS = 3f;

        final Color COLOR_PRIMARIO   = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA      = ReporteUtil.colorZebraClaro();

        final String[] HEADERS = { "PÁGINA", "FUNCIÓN", "MÓDULO", "APLICACIÓN WEB", "APLICACIÓN MÓVIL" };

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            // 1) Inicialización
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4);
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

            // Por cada rol
            if (request.getRoles() != null) {
                for (RolDTO rol : request.getRoles()) {

                    // Encabezado de rol
                    PdfPTable encabezado = new PdfPTable(1);
                    encabezado.setWidthPercentage(WIDTH_PERCENT_100);
                    encabezado.setSpacingBefore(SPACING_BEFORE_ENCABEZADO);

                    PdfPCell celdaRol = new PdfPCell(new Phrase("ROL: " + rol.getNombre(), ReporteUtil.fuenteEncabezado()));
                    celdaRol.setBackgroundColor(COLOR_PRIMARIO);
                    celdaRol.setPadding(PADDING_TITULOS);
                    encabezado.addCell(celdaRol);

                    encabezado.setSpacingAfter(SPACING_AFTER_ENCABEZADO);
                    document.add(encabezado);

                    // Subtítulo
                    PdfPTable subtitulo = new PdfPTable(1);
                    subtitulo.setWidthPercentage(WIDTH_PERCENT_100);

                    PdfPCell celdaTitulo = new PdfPCell(
                            new Phrase("FUNCIONES DEL SISTEMA ASIGNADAS", ReporteUtil.fuenteEncabezadoTablaData()));
                    celdaTitulo.setBackgroundColor(COLOR_SECUNDARIO);
                    celdaTitulo.setHorizontalAlignment(PdfPCell.ALIGN_CENTER);
                    celdaTitulo.setPadding(PADDING_TITULOS);
                    subtitulo.addCell(celdaTitulo);
                    document.add(subtitulo);

                    // Tabla de información
                    PdfPTable tabla = new PdfPTable(HEADERS.length);
                    tabla.setWidthPercentage(WIDTH_PERCENT_100);
                    tabla.setWidths(WIDTHS_INFO);
                    tabla.setSpacingBefore(SPACING_BEFORE_TABLA);

                    // Encabezados
                    for (String h : HEADERS) {
                        tabla.addCell(ReporteUtil.crearCelda(h, ReporteUtil.fuenteEncabezadoTablaData(), COLOR_SECUNDARIO));
                    }

                    // Cuerpo con zebra
                    boolean zebra = false;
                    if (rol.getFunciones() != null) {
                        for (FuncionDTO f : rol.getFunciones()) {
                            Color fondo = zebra ? COLOR_ZEBRA : Color.WHITE;
                            zebra = !zebra;

                            tabla.addCell(ReporteUtil.crearCelda(f.getPagina(), ReporteUtil.fuenteTablaData(), fondo));
                            tabla.addCell(ReporteUtil.crearCelda(f.getAccion(), ReporteUtil.fuenteTablaData(), fondo));
                            tabla.addCell(ReporteUtil.crearCelda(
                                    transformarModulo(f.getNombre_modulo()), ReporteUtil.fuenteTablaData(), fondo));

                            // Mantener la lógica original: WEB = "Sí" cuando !movil; MÓVIL = "Sí" cuando movil
                            tabla.addCell(ReporteUtil.crearCelda(f.getMovil() ? "" : "Sí", ReporteUtil.fuenteTablaData(), fondo));
                            tabla.addCell(ReporteUtil.crearCelda(f.getMovil() ? "Sí" : "", ReporteUtil.fuenteTablaData(), fondo));
                        }
                    }

                    document.add(tabla);
                }
            }

            // 3) Cierre + retorno
            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → el controller podrá responder 400
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteRoles.pdf", e);
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


    public byte[] generarReporteRolesXLSX(ReporteRolesRequest request) {
        // =========================
        // 0) Constantes DRY locales
        // =========================
        final String NOMBRE_HOJA    = "Funcionalidades de Rol"; // ≤ 31 chars
        final int    FILA_ENCABEZADO = 5;                       // fila 6 (idx 5)

        // Merges para cabecera (B1:G1 ... B5:G5) → col 1..6 (B..G)
        final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
        final int MERGE_COL_INI = 1, MERGE_COL_FIN = 6;

        final String[] HEADERS = {
            "ITEM", "ROL", "PÁGINA", "FUNCIÓN", "MÓDULO", "APLICACIÓN WEB", "APLICACIÓN MÓVIL"
        };

        final int[] ANCHOS = { 10, 30, 40, 60, 30, 20, 20 };

        // Filtros: ITEM sin filtro; resto con filtro
        final boolean[] FILTROS = new boolean[] { false, true, true, true, true, true, true };

        try (XSSFWorkbook libro = new XSSFWorkbook();
            ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1); // mantener visible encabezado

            // 1) Logo estándar A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(libro, hoja, logo); // A1:B5
            }

            // 2) MERGES exactos (B1:G1 ... B5:G5)
            for (int r = MERGE_FIL_INI; r <= MERGE_FIL_FIN; r++) {
                UtilExcel.combinarCeldas(hoja, r, r, MERGE_COL_INI, MERGE_COL_FIN);
            }

            // 3) TÍTULOS (mismo estilo que el resto)
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(
                hoja, 0, 1,
                UtilExcel.aMayusculasSeguras(request.getEmpresa()),
                estiloTitulo
            ); // B1
            UtilExcel.establecerTexto(hoja, 1, 1, "PERMISOS O FUNCIONALIDADES DEL ROL", estiloTitulo); // B2

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
            int item = 1;

            CellStyle estCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estIzqBorde    = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

            if (request.getRoles() != null) {
                for (RolDTO rol : request.getRoles()) {
                    if (rol.getFunciones() == null) continue;

                    for (FuncionDTO f : rol.getFunciones()) {
                        Row r = UtilExcel.asegurarFila(hoja, filaActual++);

                        UtilExcel.establecerValor(r, 0, item++,            estCentroBorde); // ITEM
                        UtilExcel.establecerValor(r, 1, rol.getNombre(),   estIzqBorde);    // ROL
                        UtilExcel.establecerValor(r, 2, f.getPagina(),     estIzqBorde);    // PÁGINA
                        UtilExcel.establecerValor(r, 3, f.getAccion(),     estIzqBorde);    // FUNCIÓN
                        UtilExcel.establecerValor(r, 4, transformarModulo(f.getNombre_modulo()), estIzqBorde); // MÓDULO
                        UtilExcel.establecerValor(r, 5, f.getMovil() ? "" : "Sí", estCentroBorde); // APP WEB
                        UtilExcel.establecerValor(r, 6, f.getMovil() ? "Sí" : "", estCentroBorde); // APP MÓVIL
                    }
                }
            }

            int ultimaFila = (filaActual == filaDatosInicio) ? FILA_ENCABEZADO : (filaActual - 1);

            // 6) ALINEACIONES + BORDES por regiones (como en Provincias)
            // Encabezado centrado con borde
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estCentroBorde, true);

            // Cuerpo: col 0 centrada; resto izquierda
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 0, 0, estCentroBorde, true);
                UtilExcel.aplicarEstiloARegion(hoja, filaDatosInicio, ultimaFila, 1, HEADERS.length - 1, estIzqBorde, true);
            }

            // 7) TABLA estilizada + AutoFilter (ITEM sin filtro)
            if (ultimaFila >= filaDatosInicio) {
                UtilExcel.crearTablaEstilizada(
                    hoja,
                    "RolesTabla",
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
            // Validaciones → 400 (deja pasar)
            throw e;
        } catch (Exception e) {
            // Interno → 500 uniforme
            throw new ReportBuildException("No se pudo generar Roles.xlsx", e);
        }
    }

    
    // ======================= CSV =======================
    public byte[] generarReporteRolesCSV(ReporteRolesRequest request) {
        // === Contrato del CSV ===
        final String NOMBRE_REPORTE = "Roles.csv";
        final String DELIM = ",";
        final String EOL = "\r\n"; // CRLF para Excel/Windows
        final String[] HEADERS = { "n", "rol", "pagina", "funcion", "modulo", "aplicacion_web", "aplicacion_movil" };

        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados (orden exacto)
            for (int i = 0; i < HEADERS.length; i++) {
                if (i > 0) sb.append(DELIM);
                sb.append(HEADERS[i]);
            }
            sb.append(EOL);

            // Cuerpo
            int n = 1;
            List<RolDTO> roles = request.getRoles();
            if (roles != null && !roles.isEmpty()) {
                for (RolDTO rol : roles) {
                    List<FuncionDTO> funciones = (rol == null) ? null : rol.getFunciones();
                    if (funciones == null || funciones.isEmpty()) continue;

                    for (FuncionDTO f : funciones) {
                        String nombreRol = (rol == null || rol.getNombre() == null) ? "" : rol.getNombre();

                        String pagina = (f == null || f.getPagina() == null) ? "" : f.getPagina();
                        String accion = (f == null || f.getAccion() == null) ? "" : f.getAccion();

                        String modulo = "";
                        if (f != null) {
                            String nomModulo = f.getNombre_modulo();
                            modulo = (nomModulo == null) ? "" : transformarModulo(nomModulo);
                        }

                        boolean esMovil = (f != null && ((f.getMovil() instanceof Boolean) ? Boolean.TRUE.equals(f.getMovil()) : false));
                        String appWeb   = esMovil ? ""  : "Sí";
                        String appMovil = esMovil ? "Sí" : "";

                        sb.append(n++).append(DELIM)
                        .append(UtilCsv.csvEscape(nombreRol)).append(DELIM)
                        .append(UtilCsv.csvEscape(pagina)).append(DELIM)
                        .append(UtilCsv.csvEscape(accion)).append(DELIM)
                        .append(UtilCsv.csvEscape(modulo)).append(DELIM)
                        .append(UtilCsv.csvEscape(appWeb)).append(DELIM)
                        .append(UtilCsv.csvEscape(appMovil)).append(EOL);
                    }
                }
            }

            // Retorno (nunca null)
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);

        } catch (IllegalArgumentException e) {
            // Validación → 400
            throw e;
        } catch (Exception e) {
            // Interno → 500
            throw new ReportBuildException("No se pudo generar " + NOMBRE_REPORTE, e);
        }
    }

        
    // ======================= XML =======================
    public byte[] generarReporteRolesXML(ReporteRolesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();
            sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            sb.append("<Roles>\n");

            if (request.getRoles() != null) {
                for (RolDTO rol : request.getRoles()) {
                    sb.append("  <rol");
                    // Si tu RolDTO tiene getId(), descomenta la siguiente línea:
                    // sb.append(" id=\"").append(xmlEsc(String.valueOf(rol.getId()))).append("\"");
                    sb.append(">\n");

                    sb.append("    <nombre>").append(xmlEsc(rol.getNombre())).append("</nombre>\n");
                    sb.append("    <funciones>\n");

                    if (rol.getFunciones() != null) {
                        for (FuncionDTO f : rol.getFunciones()) {
                            String modulo = transformarModulo(f.getNombre_modulo());
                            String appWeb = f.getMovil() ? "" : "Sí";
                            String appMovil = f.getMovil() ? "Sí" : "";

                            sb.append("      <detalle>\n");
                            sb.append("        <pagina>").append(xmlEsc(f.getPagina())).append("</pagina>\n");
                            sb.append("        <funcion>").append(xmlEsc(f.getAccion())).append("</funcion>\n");
                            sb.append("        <modulo>").append(xmlEsc(modulo)).append("</modulo>\n");
                            sb.append("        <aplicacion_web>").append(xmlEsc(appWeb)).append("</aplicacion_web>\n");
                            sb.append("        <aplicacion_movil>").append(xmlEsc(appMovil))
                                    .append("</aplicacion_movil>\n");
                            sb.append("      </detalle>\n");
                        }
                    }

                    sb.append("    </funciones>\n");
                    sb.append("  </rol>\n");
                }
            }

            sb.append("</Roles>\n");
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private String xmlEsc(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }


    // METODO AUXILIAR PARA CONVERTIR EL DATO A UN TEXTO MAS AMIGABLE
    private String transformarModulo(String nombreModulo) {
        if (nombreModulo == null)
            return "";
        switch (nombreModulo) {
            case "permisos":
                return "Módulo de Permisos";
            case "vacaciones":
                return "Módulo de Vacaciones";
            case "horas_extras":
                return "Módulo de Horas Extras";
            case "alimentacion":
                return "Módulo de Alimentación";
            case "acciones_personal":
                return "Módulo de Acciones de Personal";
            case "geolocalizacion":
                return "Módulo de Geolocalización";
            case "timbre_virtual":
                return "Módulo de Timbre Virtual";
            case "reloj_virtual":
                return "Aplicación Móvil";
            case "aprobar":
                return "Aprobaciones Solicitudes";
            default:
                return nombreModulo;
        }
    }

}
