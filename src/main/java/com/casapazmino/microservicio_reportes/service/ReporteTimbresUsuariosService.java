package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresUsuarios.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;

import org.springframework.stereotype.Service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.common.usermodel.HyperlinkType;

import java.awt.Color;
import java.io.OutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteTimbresUsuariosService {

    /* ========================== PDF (STREAMING) ========================== */
    public void escribirReportePDF(ReporteTimbresUsuariosRequest request, OutputStream out) throws Exception {

        // ➊ DRY: constantes locales (look & feel intacto)
        final boolean LANDSCAPE = Boolean.TRUE.equals(request.getTimbreDispositivo());
        final Rectangle PAGE = LANDSCAPE ? PageSize.A4.rotate() : PageSize.A4;

        final String TITULO = "TIMBRES - " + (request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
        final String PERIODO = "PERIODO DEL: " + request.getPeriodo().getInicio() + " AL "
                + request.getPeriodo().getFin();

        final float[] WIDTHS_TITULO = { 8f, 2f };
        final float[] WIDTHS_INFO = { 4f, 4f, 4f };
        final float[] WIDTHS_TIMBRES = { 1f, 2.5f, 1.5f, 1.5f, 2.3f, 4.5f, 2f, 2f };
        final int WIDTH_PERCENT_100 = 100;
        final float SPACING_AFTER_TIT = 10f;
        final float PADDING_TITULOS = 5f;

        final Color COLOR_PRIMARIO = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer = null;

        try {
            // 1) Inicialización de recursos PDF
            document = new Document(PAGE, 40, 40, 30, 50);
            writer = PdfWriter.getInstance(document, out);
            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    request.getUsuario(),
                    request.getFraseMarcaAgua(),
                    request.getColorPrincipal()));
            document.open();

            // 2) Construcción (solo invocando helpers existentes)

            // Logo
            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null)
                document.add(logo);

            // Encabezados
            document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
            document.add(ReporteUtil.crearTituloReporte(TITULO));
            document.add(ReporteUtil.crearTituloPeriodo(PERIODO));

            // Contador global
            AtomicInteger contadorGlobal = new AtomicInteger();
            if (request.getData_pdf() != null) {
                request.getData_pdf().forEach(g -> {
                    if (g.getEmpleados() != null) {
                        g.getEmpleados().forEach(
                                e -> contadorGlobal.addAndGet(e.getTimbres() != null ? e.getTimbres().size() : 0));
                    }
                });
            }

            // Título de bloque + contador
            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(WIDTH_PERCENT_100);
            tituloTabla.setWidths(WIDTHS_TITULO);
            tituloTabla.setSpacingAfter(SPACING_AFTER_TIT);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezadoTablaData()));
            celdaTitulo.setBackgroundColor(COLOR_SECUNDARIO);
            celdaTitulo.setPadding(PADDING_TITULOS);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase("N° Registros: " + contadorGlobal.get(), ReporteUtil.fuenteEncabezadoTablaData()));
            celdaContador.setBackgroundColor(COLOR_SECUNDARIO);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(PADDING_TITULOS);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);

            document.add(tituloTabla);

            // Grupos / Empleados / Timbres
            if (request.getData_pdf() != null) {
                for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                    if (grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoTimbreDTO emp : grupo.getEmpleados()) {

                        // Ficha empleado con borde exterior
                        PdfPTable infoEmpleado = new PdfPTable(3);
                        infoEmpleado.setWidthPercentage(WIDTH_PERCENT_100);
                        infoEmpleado.setWidths(WIDTHS_INFO);

                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                                emp.getApellido() + " " + emp.getNombre(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), COLOR_ZEBRA));
                        infoEmpleado
                                .addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), COLOR_ZEBRA));
                        infoEmpleado.addCell(
                                ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), COLOR_ZEBRA));
                        infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), COLOR_ZEBRA));

                        PdfPTable tablaContenedora = new PdfPTable(1);
                        tablaContenedora.setWidthPercentage(WIDTH_PERCENT_100);
                        PdfPCell contenedor = new PdfPCell(infoEmpleado);
                        contenedor.setPadding(0);
                        contenedor.setBorder(Rectangle.BOX);
                        tablaContenedora.addCell(contenedor);
                        document.add(tablaContenedora);

                        // Tabla de timbres (8 columnas, encabezado 2 filas)
                        PdfPTable tablaTimbres = new PdfPTable(8);
                        tablaTimbres.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaTimbres.setSpacingBefore(0f);
                        tablaTimbres.setWidths(WIDTHS_TIMBRES);

                        // Encabezado fila 1
                        PdfPCell celdaN = ReporteUtil.crearCelda("N°", ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO);
                        celdaN.setRowspan(2);
                        celdaN.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaN.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaN);

                        PdfPCell celdaTimbre = ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);
                        celdaTimbre.setColspan(2);
                        celdaTimbre.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaTimbre.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaTimbre);

                        PdfPCell celdaReloj = ReporteUtil.crearCelda("RELOJ", ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);
                        celdaReloj.setRowspan(2);
                        celdaReloj.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaReloj.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaReloj);

                        PdfPCell celdaAccion = ReporteUtil.crearCelda("ACCIÓN", ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);
                        celdaAccion.setRowspan(2);
                        celdaAccion.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaAccion.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaAccion);

                        PdfPCell celdaObs = ReporteUtil.crearCelda("OBSERVACIÓN", ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);
                        celdaObs.setRowspan(2);
                        celdaObs.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaObs.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaObs);

                        PdfPCell celdaUbicacion = ReporteUtil.crearCelda(
                                "UBICACIÓN",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);
                        celdaUbicacion.setRowspan(2);
                        celdaUbicacion.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaUbicacion.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaUbicacion);

                        PdfPCell celdaImagen = ReporteUtil.crearCelda(
                                "IMAGEN",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO);
                        celdaImagen.setRowspan(2);
                        celdaImagen.setHorizontalAlignment(Element.ALIGN_CENTER);
                        celdaImagen.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        tablaTimbres.addCell(celdaImagen);

                        // Encabezado fila 2
                        tablaTimbres.addCell(
                                ReporteUtil.crearCelda("FECHA", ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));
                        tablaTimbres.addCell(
                                ReporteUtil.crearCelda("HORA", ReporteUtil.fuenteEncabezadoTablaData(), COLOR_PRIMARIO));

                        // Cuerpo (zebra)
                        int contadorLocal = 1;
                        if (emp.getTimbres() != null) {
                            for (TimbreUsuarioDTO t : emp.getTimbres()) {
                                Color fondo = (contadorLocal % 2 == 0) ? COLOR_ZEBRA : Color.WHITE;

                                String[] partes = t.getFecha_hora_timbre().split(" ");
                                String fecha = partes.length > 0 ? partes[0] : "";
                                String hora = partes.length > 1 ? partes[1] : "";

                                tablaTimbres.addCell(ReporteUtil.crearCelda(String.valueOf(contadorLocal),
                                        ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.formatearFechaConDia(fecha),
                                        ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(hora, ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(
                                        ReporteUtil.crearCelda(t.getId_reloj(), ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(ReporteUtil.crearCelda(ReporteUtil.traducirAccion(t.getAccion()),
                                        ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(
                                        ReporteUtil.crearCelda(t.getObservacion(), ReporteUtil.fuenteTablaData(), fondo));
                                tablaTimbres.addCell(
                                        crearCeldaUbicacionPDF(
                                                t.getLatitud(),
                                                t.getLongitud(),
                                                fondo));

                                tablaTimbres.addCell(
                                        crearCeldaImagenPDF(
                                                t.getImagenUrl(),
                                                fondo));
                                contadorLocal++;
                            }
                        }

                        document.add(tablaTimbres);
                        document.add(Chunk.NEWLINE);
                    }
                }
            }

            // 3) Cierre (el OutputStream 'out' lo cierra el contenedor/llamador)
            document.close();

        } catch (IllegalArgumentException e) {
            // Validaciones de helpers → dejar pasar (el controller podrá mapear a 400)
            throw e;
        } catch (Exception e) {
            // Fallo interno uniforme → 500
            throw new ReportBuildException("No se pudo generar ReporteTimbresUsuarios.pdf", e);
        } finally {
            // 4) Ciclo de recursos garantizado (sin cerrar 'out')
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
            // NO cerrar 'out' aquí
        }
    }

    /*
     * ========================== XLSX (STREAMING SIEMPRE)
     * ==========================
     */
    public void escribirReporteTimbresUsuariosExcel(ReporteTimbresUsuariosRequest request, OutputStream out)
            throws Exception {

        final boolean conDispositivo = request.getTimbreDispositivo();

        // 1) Definir columnas según el flag (igual que antes)
        String[] headers = conDispositivo
                ? new String[] { "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL", "RÉGIMEN",
                        "DEPARTAMENTO", "CARGO", "FECHA TIMBRE", "HORA TIMBRE", "RELOJ", "ACCIÓN", "OBSERVACIÓN",
                        "UBICACIÓN", "IMAGEN", "FECHA TIMBRE DISPOSITIVO", "HORA TIMBRE DISPOSITIVO" }
                : new String[] { "ITEM", "IDENTIFICACIÓN", "CÓDIGO", "APELLIDO NOMBRE", "CIUDAD", "SUCURSAL", "RÉGIMEN",
                        "DEPARTAMENTO", "CARGO", "FECHA TIMBRE", "HORA TIMBRE", "RELOJ", "ACCIÓN", "OBSERVACIÓN",
                        "UBICACIÓN", "IMAGEN" };
        int lastCol = headers.length - 1;

        // 2) Libro base + streaming (cerrando ambos recursos)
        try (XSSFWorkbook wbBase = new XSSFWorkbook();
                SXSSFWorkbook wb = new SXSSFWorkbook(wbBase, 200, true, false)) {

            Sheet hoja = wb.createSheet("Timbres");
            // Fijar encabezado (fila 6 → idx 5)
            hoja.createFreezePane(0, 6);

            // === ESTILOS (igual base, reusables) ===
            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(wbBase);

            XSSFColor azulCorporativo = new XSSFColor(new java.awt.Color(14, 91, 166), null);
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(wbBase);
            if (estiloHeader instanceof org.apache.poi.xssf.usermodel.XSSFCellStyle) {
                org.apache.poi.xssf.usermodel.XSSFCellStyle x = (org.apache.poi.xssf.usermodel.XSSFCellStyle) estiloHeader;
                x.setFillForegroundColor(
                        new org.apache.poi.xssf.usermodel.XSSFColor(hexToAwtColor(safe(azulCorporativo)), null));
                x.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else {
                estiloHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            Font fHeader = wbBase.createFont();
            fHeader.setBold(true);
            fHeader.setColor(IndexedColors.WHITE.getIndex());
            estiloHeader.setFont(fHeader);
            estiloHeader.setWrapText(true);

            // Bordes para cuerpo (base)
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(wbBase);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(wbBase);

            // Zebra (suave)
            CellStyle estiloCentroZebra = wbBase.createCellStyle();
            estiloCentroZebra.cloneStyleFrom(estiloCentroBorde);
            estiloCentroZebra.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            estiloCentroZebra.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());

            CellStyle estiloIzqZebra = wbBase.createCellStyle();
            estiloIzqZebra.cloneStyleFrom(estiloIzqBorde);
            estiloIzqZebra.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            estiloIzqZebra.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());

            // === ESTILOS PARA HIPERVÍNCULOS ===
            org.apache.poi.ss.usermodel.Font fuenteLink = wbBase.createFont();
            fuenteLink.setColor(IndexedColors.BLUE.getIndex());
            fuenteLink.setUnderline(org.apache.poi.ss.usermodel.Font.U_SINGLE);

            // Link fondo blanco
            CellStyle estiloLinkCentro = wbBase.createCellStyle();
            estiloLinkCentro.cloneStyleFrom(estiloCentroBorde);
            estiloLinkCentro.setFont(fuenteLink);

            // Link fondo zebra
            CellStyle estiloLinkCentroZebra = wbBase.createCellStyle();
            estiloLinkCentroZebra.cloneStyleFrom(estiloCentroZebra);
            estiloLinkCentroZebra.setFont(fuenteLink);


            // === LOGO A1:B5 ===
            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                UtilExcel.insertarLogoEstandar(wb, hoja, logo);
            }

            // === MERGES B1..B5 hasta lastCol ===
            for (int row = 0; row <= 4; row++) {
                UtilExcel.combinarCeldas(hoja, row, row, 1, lastCol);
            }

            // === TÍTULOS (B1..B3) ===
            String empresa = UtilExcel.aMayusculasSeguras(request.getEmpresa());
            String activosInactivos = ("1".equals(String.valueOf(request.getOpcionBusqueda()))) ? "ACTIVOS"
                    : "INACTIVOS";
            String periodo = "PERIODO DEL REPORTE: " +
                    (request.getPeriodo() != null ? safe(request.getPeriodo()::getInicio) : "") +
                    " AL " +
                    (request.getPeriodo() != null ? safe(request.getPeriodo()::getFin) : "");

            UtilExcel.establecerTexto(hoja, 0, 1, empresa, estiloTitulo);
            hoja.getRow(0).setHeightInPoints(22f);
            UtilExcel.establecerTexto(hoja, 1, 1, ("LISTA DE TIMBRES - " + activosInactivos).toUpperCase(),
                    estiloTitulo);
            hoja.getRow(1).setHeightInPoints(20f);
            UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);
            hoja.getRow(2).setHeightInPoints(18f);
            UtilExcel.asegurarFila(hoja, 3); // blanco

            // === ENCABEZADOS (fila 6 → idx 5) ===
            final int filaEnc = 5;
            Row filaHeader = UtilExcel.asegurarFila(hoja, filaEnc);
            filaHeader.setHeightInPoints(18f);
            for (int c = 0; c < headers.length; c++) {
                UtilExcel.establecerTexto(filaHeader, c, headers[c], estiloHeader);
            }

            // Anchos (criterios del original)
            int[] anchos = new int[headers.length];
            for (int c = 0; c < headers.length; c++) {
                int width;
                switch (c) {
                    case 0:
                        width = 10;
                        break; // ITEM
                    case 3:
                        width = 28;
                        break; // APELLIDO NOMBRE
                    case 13:
                        width = 26;
                        break; // OBSERVACIÓN
                    case 16:
                        width = 28;
                        break; // FECHA TIMBRE DISPOSITIVO
                    case 17:
                        width = 24;
                        break; // HORA TIMBRE DISPOSITIVO
                    default:
                        width = 20;
                }
                anchos[c] = width;
            }
            UtilExcel.establecerAnchosColumnas(hoja, anchos);

            // === DATOS (aplicando estilos celda-a-celda con cebra) ===
            int rowIdx = filaEnc + 1;
            int item = 1;

            if (request.getData_pdf() != null) {
                for (GrupoTimbresDTO grupo : request.getData_pdf()) {
                    if (grupo == null || grupo.getEmpleados() == null)
                        continue;

                    for (EmpleadoTimbreDTO emp : grupo.getEmpleados()) {
                        if (emp == null || emp.getTimbres() == null)
                            continue;

                        String apenom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();
                        String ciudad = firstNonEmpty(safe(emp.getCiudad()), safe(grupo.getCiudad()));
                        String sucursal = firstNonEmpty(safe(emp.getSucursal()), safe(grupo.getSucursal()));

                        for (TimbreUsuarioDTO t : emp.getTimbres()) {
                            if (t == null)
                                continue;

                            String fhVal = safe(t.getFecha_hora_timbre_validado());
                            String fechaSrv = fechaDeFechaHora(fhVal);
                            String horaSrv = horaDeFechaHora(fhVal);

                            String fhDev = safe(t.getFecha_hora_timbre());
                            String fechaDev = fechaDeFechaHora(fhDev);
                            String horaDev = horaDeFechaHora(fhDev);

                            String fechaFmt = ReporteUtil
                                    .formatearFechaConDia(fechaSrv.isBlank() ? fechaDev : fechaSrv);
                            String hora = horaSrv.isBlank() ? horaDev : horaSrv;
                            String accion = ReporteUtil.traducirAccion(safe(t.getAccion()));

                            Row r = hoja.createRow(rowIdx++);
                            int col = 0;

                            boolean zebra = ((rowIdx - (filaEnc + 1)) % 2) == 1;
                            CellStyle csCenter = zebra ? estiloCentroZebra : estiloCentroBorde;
                            CellStyle csLeft = zebra ? estiloIzqZebra : estiloIzqBorde;

                            Cell c0 = r.createCell(col++);
                            c0.setCellValue(item++);
                            c0.setCellStyle(csCenter);
                            Cell c1 = r.createCell(col++);
                            c1.setCellValue(safe(emp.getIdentificacion()));
                            c1.setCellStyle(csLeft);
                            Cell c2 = r.createCell(col++);
                            c2.setCellValue(safe(emp.getCodigo()));
                            c2.setCellStyle(csLeft);
                            Cell c3 = r.createCell(col++);
                            c3.setCellValue(apenom);
                            c3.setCellStyle(csLeft);
                            Cell c4 = r.createCell(col++);
                            c4.setCellValue(ciudad);
                            c4.setCellStyle(csLeft);
                            Cell c5 = r.createCell(col++);
                            c5.setCellValue(sucursal);
                            c5.setCellStyle(csLeft);
                            Cell c6 = r.createCell(col++);
                            c6.setCellValue(safe(emp.getRegimen()));
                            c6.setCellStyle(csLeft);
                            Cell c7 = r.createCell(col++);
                            c7.setCellValue(safe(emp.getDepartamento()));
                            c7.setCellStyle(csLeft);
                            Cell c8 = r.createCell(col++);
                            c8.setCellValue(safe(emp.getCargo()));
                            c8.setCellStyle(csLeft);
                            Cell c9 = r.createCell(col++);
                            c9.setCellValue(fechaFmt);
                            c9.setCellStyle(csLeft);
                            Cell c10 = r.createCell(col++);
                            c10.setCellValue(hora);
                            c10.setCellStyle(csLeft);
                            Cell c11 = r.createCell(col++);
                            c11.setCellValue(safe(t.getId_reloj()));
                            c11.setCellStyle(csLeft);
                            Cell c12 = r.createCell(col++);
                            c12.setCellValue(accion);
                            c12.setCellStyle(csLeft);
                            Cell c13 = r.createCell(col++);
                            c13.setCellValue(safe(t.getObservacion()));
                            c13.setCellStyle(csLeft);

                            Cell c14 = r.createCell(col++);

                            establecerUbicacionExcel(
                                    wb,
                                    c14,
                                    t.getLatitud(),
                                    t.getLongitud()
                            );

                            c14.setCellStyle(
                                    c14.getHyperlink() != null
                                            ? (zebra ? estiloLinkCentroZebra : estiloLinkCentro)
                                            : csCenter
                            );


                            Cell c15 = r.createCell(col++);

                            establecerImagenExcel(
                                    wb,
                                    c15,
                                    t.getImagenUrl()
                            );

                            c15.setCellStyle(
                                    c15.getHyperlink() != null
                                            ? (zebra ? estiloLinkCentroZebra : estiloLinkCentro)
                                            : csCenter
                            );

                            if (conDispositivo) {
                                Cell c16 = r.createCell(col++);
                                c16.setCellValue(fechaDev);
                                c16.setCellStyle(csLeft);
                                Cell c17 = r.createCell(col++);
                                c17.setCellValue(horaDev);
                                c17.setCellStyle(csLeft);
                            }
                        }
                    }
                }
            }

            // === AutoFilter (sin UtilExcel.crearTablaEstilizada; evita conflicto
            // XSSF/SXSSF) ===
            if (rowIdx > filaEnc + 1) {
                hoja.setAutoFilter(new CellRangeAddress(filaEnc, rowIdx - 1, 0, lastCol));
            }

            // === Escribir en el OutputStream ===
            wb.write(out);
            wb.dispose();

        } catch (IllegalArgumentException e) {
            // Validación → que el controller mapee a 400
            throw e;
        } catch (Exception e) {
            // Internos → 500 uniforme
            throw new ReportBuildException("No se pudo generar TimbresUsuarios.xlsx", e);
        }
    }

    /* ========================== Helpers ========================== */
    private String safe(Object v) {
        if (v == null)
            return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private String safe(java.util.concurrent.Callable<String> c) {
        try {
            String s = c.call();
            return s == null ? "" : s;
        } catch (Exception e) {
            return "";
        }
    }

    private String firstNonEmpty(String a, String b) {
        return (a == null || a.isBlank()) ? (b == null ? "" : b) : a;
    }

    private String fechaDeFechaHora(String fh) {
        if (fh == null || fh.isBlank())
            return "";
        int i = fh.indexOf(' ');
        return (i > 0) ? fh.substring(0, i) : fh;
    }

    private String horaDeFechaHora(String fh) {
        if (fh == null || fh.isBlank())
            return "";
        int i = fh.indexOf(' ');
        return (i >= 0 && i + 1 < fh.length()) ? fh.substring(i + 1) : "";
    }

    private java.awt.Color hexToAwtColor(String hex) {
        if (hex == null)
            return new java.awt.Color(0, 112, 192); // azul por defecto
        String s = hex.trim().replace("#", "");
        try {
            int r = Integer.valueOf(s.substring(0, 2), 16);
            int g = Integer.valueOf(s.substring(2, 4), 16);
            int b = Integer.valueOf(s.substring(4, 6), 16);
            return new java.awt.Color(r, g, b);
        } catch (Exception e) {
            return new java.awt.Color(0, 112, 192);
        }
    }


    private PdfPCell crearCeldaUbicacionPDF(
            String latitud,
            String longitud,
            Color fondo) {

        String lat = safe(latitud);
        String lon = safe(longitud);

        PdfPCell celda = ReporteUtil.crearCelda(
                "",
                ReporteUtil.fuenteTablaData(),
                fondo
        );

        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);

        if (lat.isBlank() || lon.isBlank()) {
            return celda;
        }

        String url = "https://www.google.com/maps/search/?api=1&query="
                + lat + "," + lon;

        org.openpdf.text.Font fuenteBase =
                ReporteUtil.fuenteTablaData();

        org.openpdf.text.Font fuenteLink =
                new org.openpdf.text.Font(
                        fuenteBase.getFamily(),
                        fuenteBase.getSize(),
                        org.openpdf.text.Font.UNDERLINE,
                        Color.BLUE
                );

        Chunk enlace = new Chunk(
                "Ubicación",
                fuenteLink
        );

        enlace.setAnchor(url);

        celda.setPhrase(new Phrase(enlace));

        return celda;
    }

    private PdfPCell crearCeldaImagenPDF(
            String imagenUrl,
            Color fondo) {

        PdfPCell celda = ReporteUtil.crearCelda(
                "",
                ReporteUtil.fuenteTablaData(),
                fondo
        );

        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);

        String url = safe(imagenUrl);

        if (url.isBlank()) {
            return celda;
        }

        org.openpdf.text.Font fuenteBase =
                ReporteUtil.fuenteTablaData();

        org.openpdf.text.Font fuenteLink =
                new org.openpdf.text.Font(
                        fuenteBase.getFamily(),
                        fuenteBase.getSize(),
                        org.openpdf.text.Font.UNDERLINE,
                        Color.BLUE
                );

        Chunk enlace = new Chunk(
                "Imagen",
                fuenteLink
        );

        enlace.setAnchor(url);

        celda.setPhrase(new Phrase(enlace));

        return celda;
    }
        

    private void establecerUbicacionExcel(
            Workbook workbook,
            Cell celda,
            String latitud,
            String longitud) {

        String lat = safe(latitud);
        String lon = safe(longitud);

        if (lat.isBlank() || lon.isBlank()) {
            celda.setCellValue("");
            return;
        }

        String url = "https://www.google.com/maps/search/?api=1&query="
                + lat + "," + lon;

        CreationHelper helper = workbook.getCreationHelper();

        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(url);

        celda.setCellValue("Ubicación");
        celda.setHyperlink(hyperlink);
    }

    private void establecerImagenExcel(
            Workbook workbook,
            Cell celda,
            String imagenUrl) {

        String url = safe(imagenUrl);

        if (url.isBlank()) {
            celda.setCellValue("");
            return;
        }

        CreationHelper helper = workbook.getCreationHelper();

        Hyperlink hyperlink = helper.createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress(url);

        celda.setCellValue("Imagen");
        celda.setHyperlink(hyperlink);
    }

}
