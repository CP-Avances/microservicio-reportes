package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.Rol.FuncionDTO;
import com.casapazmino.microservicio_reportes.model.Rol.ReporteRolesRequest;
import com.casapazmino.microservicio_reportes.model.Rol.RolDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
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
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet hoja = wb.createSheet("Funcionalidades de Rol");

            // 1) Logo A1:B5
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            UtilExcel.insertarLogoEstandar(wb, hoja, logo);

            // 2) Estilos
            CellStyle estTitulo = ConfiguracionExcel.crearEstiloTitulo(wb);
            CellStyle estEnc = ConfiguracionExcel.crearEstiloEncabezadoTabla(wb);
            CellStyle estC = ConfiguracionExcel.crearEstiloCentroConBorde(wb);
            CellStyle estI = ConfiguracionExcel.crearEstiloIzquierdaConBorde(wb);

            // 3) Títulos (B1:G1 y B2:G2)
            UtilExcel.combinarCeldas(hoja, 0, 0, 1, 6); // B1:G1
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estTitulo);

            UtilExcel.combinarCeldas(hoja, 1, 1, 1, 6); // B2:G2
            UtilExcel.establecerTexto(hoja, 1, 1, "PERMISOS O FUNCIONALIDADES DEL ROL", estTitulo);

            // 4) Encabezados (fila 6 visual -> índice 5)
            final int filaEnc = 5;
            String[] headers = { "ITEM", "ROL", "PÁGINA", "FUNCIÓN", "MÓDULO", "APLICACIÓN WEB", "APLICACIÓN MÓVIL" };
            for (int c = 0; c < headers.length; c++) {
                UtilExcel.establecerTexto(hoja, filaEnc, c, headers[c], estEnc); // esta sobrecarga sí acepta
                                                                                 // Sheet+índices
            }
            // aplicarEstiloAFila espera Row + #columnas
            Row filaEncRow = UtilExcel.asegurarFila(hoja, filaEnc);
            UtilExcel.aplicarEstiloAFila(filaEncRow, headers.length, estEnc);

            // 5) Datos (establecerValor espera Row)
            int fila = filaEnc + 1;
            int item = 1;
            if (request.getRoles() != null) {
                for (RolDTO rol : request.getRoles()) {
                    if (rol.getFunciones() == null)
                        continue;
                    for (FuncionDTO f : rol.getFunciones()) {
                        Row r = UtilExcel.asegurarFila(hoja, fila);

                        UtilExcel.establecerValor(r, 0, item++, estC); // ITEM
                        UtilExcel.establecerValor(r, 1, rol.getNombre(), estI); // ROL
                        UtilExcel.establecerValor(r, 2, f.getPagina(), estI); // PÁGINA
                        UtilExcel.establecerValor(r, 3, f.getAccion(), estI); // FUNCIÓN
                        UtilExcel.establecerValor(r, 4, transformarModulo(f.getNombre_modulo()), estI); // MÓDULO
                        UtilExcel.establecerValor(r, 5, f.getMovil() ? "" : "Sí", estC); // APP WEB
                        UtilExcel.establecerValor(r, 6, f.getMovil() ? "Sí" : "", estC); // APP MÓVIL

                        fila++;
                    }
                }
            }

            // 6) Tabla visual (firma real: ... , boolean mostrarRayadoFilas, boolean[]
            // filtroPorColumna)
            UtilExcel.crearTablaEstilizada(
                    hoja,
                    "RolesTabla",
                    filaEnc, 0,
                    Math.max(fila - 1, filaEnc), 6,
                    true, // mostrar rayado (zebra)
                    null // filtros por columna (no usado en esta versión)
            );

            // 7) Anchos
            UtilExcel.establecerAnchosColumnas(hoja, new int[] { 10, 30, 40, 60, 30, 20, 20 });

            // 8) Salida
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            return out.toByteArray();

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    // ======================= CSV =======================
    public byte[] generarReporteRolesCSV(ReporteRolesRequest request) {
        try {
            StringBuilder sb = new StringBuilder();

            // Encabezados EXACTOS como en el front antiguo
            sb.append("n,rol,pagina,funcion,modulo,aplicacion_web,aplicacion_movil\n");

            int n = 1;
            if (request.getRoles() != null) {
                for (RolDTO rol : request.getRoles()) {
                    if (rol.getFunciones() == null)
                        continue;

                    for (FuncionDTO f : rol.getFunciones()) {
                        String modulo = transformarModulo(f.getNombre_modulo());
                        String appWeb = f.getMovil() ? "" : "Sí";
                        String appMovil = f.getMovil() ? "Sí" : "";

                        sb.append(n++).append(',')
                                .append(csvEsc(rol.getNombre())).append(',')
                                .append(csvEsc(f.getPagina())).append(',')
                                .append(csvEsc(f.getAccion())).append(',')
                                .append(csvEsc(modulo)).append(',')
                                .append(csvEsc(appWeb)).append(',')
                                .append(csvEsc(appMovil)).append('\n');
                    }
                }
            }
            return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
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

    private String csvEsc(String v) {
        if (v == null)
            return "";
        boolean mustQuote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return mustQuote ? "\"" + s + "\"" : s;
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
