package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.ReporteSolicitudVacacionRequest;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.SolicitudVacacionEmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.SolicitudVacacionReporteDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.Paragraph;
import org.openpdf.text.Phrase;
import org.openpdf.text.Rectangle;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class ReporteConsolidadoSolicitudVacacionService {

    // =========================================================================================
    // PDF
    // =========================================================================================

    public byte[] generarReporteSolicitudesVacacionesPDF(ReporteSolicitudVacacionRequest request) {
        final float[] W_EMP_INFO = {4f, 4f, 2f};
        final float[] W_EMP_INFO_2 = {4f, 4f, 3f};
        final float[] W_SOL = {1.2f, 1.8f, 1.8f, 1.2f, 1.2f, 1.5f, 1.8f, 2.6f, 2.2f};

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 30, 30, 30, 50);
            writer = PdfWriter.getInstance(document, baos);

            writer.setPageEvent(new ConfiguracionPaginaPDF(
                    safe(request.getUsuario()),
                    safe(request.getFraseMarcaAgua()),
                    safe(request.getColorPrincipal())
            ));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            document.add(ReporteUtil.crearTituloEmpresa(safe(request.getEmpresa())));
            document.add(ReporteUtil.crearTituloReporte("REPORTE - SOLICITUDES DE VACACIONES"));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta())
            ));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color zebra = ReporteUtil.colorZebraClaro();

            List<SolicitudVacacionEmpleadoDTO> empleados = agruparPorEmpleado(request.getSolicitudes());

            if (empleados.isEmpty()) {
                Paragraph p = new Paragraph("Sin datos para mostrar", ReporteUtil.fuenteTexto());
                p.setAlignment(Element.ALIGN_CENTER);
                p.setSpacingBefore(20f);
                document.add(p);
                document.close();
                return baos.toByteArray();
            }

            for (SolicitudVacacionEmpleadoDTO emp : empleados) {
                if (emp == null) continue;

                // =============================================================================
                // Información del empleado
                // Se mantienen los mismos campos y distribución.
                // Únicamente se aplica el diseño estándar de los demás reportes.
                // =============================================================================

                PdfPTable infoEmp = new PdfPTable(3);
                infoEmp.setWidthPercentage(100);
                infoEmp.setWidths(W_EMP_INFO);
                infoEmp.setSpacingBefore(0f);
                infoEmp.setSpacingAfter(0f);

                infoEmp.addCell(ReporteUtil.celdaInfoMixta("CIUDAD:", safe(emp.getCiudad()), zebra));
                infoEmp.addCell(ReporteUtil.celdaInfoMixta("C.C.:", safe(emp.getIdentificacion()), zebra));
                infoEmp.addCell(ReporteUtil.celdaInfoMixta("COD:", safe(emp.getCodigo()), zebra));

                PdfPCell cNombre = ReporteUtil.celdaInfoMixta("EMPLEADO:", safe(emp.getEmpleado()), zebra);
                cNombre.setColspan(3);
                infoEmp.addCell(cNombre);

                PdfPTable infoEmp2 = new PdfPTable(3);
                infoEmp2.setWidthPercentage(100);
                infoEmp2.setWidths(W_EMP_INFO_2);
                infoEmp2.setSpacingBefore(0f);
                infoEmp2.setSpacingAfter(0f);

                infoEmp2.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", safe(emp.getRegimen()), zebra));
                infoEmp2.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", safe(emp.getDepartamento()), zebra));
                infoEmp2.addCell(ReporteUtil.celdaInfoMixta("CARGO:", safe(emp.getCargo()), zebra));

                PdfPTable contInfo = new PdfPTable(1);
                contInfo.setWidthPercentage(100);
                contInfo.setSpacingBefore(6f);
                contInfo.setSpacingAfter(0f);

                PdfPCell wrap1 = new PdfPCell(infoEmp);
                wrap1.setPadding(0f);
                wrap1.setBorder(Rectangle.TOP | Rectangle.LEFT | Rectangle.RIGHT);
                contInfo.addCell(wrap1);

                PdfPCell wrap2 = new PdfPCell(infoEmp2);
                wrap2.setPadding(0f);
                wrap2.setBorder(Rectangle.BOTTOM | Rectangle.LEFT | Rectangle.RIGHT);
                contInfo.addCell(wrap2);

                document.add(contInfo);

                // =============================================================================
                // Tabla de solicitudes
                // Se conservan exactamente las mismas nueve columnas.
                // =============================================================================

                PdfPTable tabla = new PdfPTable(9);
                tabla.setWidthPercentage(100);
                tabla.setWidths(W_SOL);
                tabla.setSpacingBefore(0f);

                tabla.addCell(hCell("SOLICITUD", colorPrincipal));
                tabla.addCell(hCell("DESDE", colorPrincipal));
                tabla.addCell(hCell("HASTA", colorPrincipal));
                tabla.addCell(hCell("DÍAS L-V", colorPrincipal));
                tabla.addCell(hCell("DÍAS S-D", colorPrincipal));
                tabla.addCell(hCell("AUTORIZADO", colorPrincipal));
                tabla.addCell(hCell("ESTADO", colorPrincipal));
                tabla.addCell(hCell("AUTORIZA", colorPrincipal));
                tabla.addCell(hCell("FECHA DE AUTORIZACIÓN", colorPrincipal));

                List<SolicitudVacacionReporteDTO> solicitudes = ordenarSolicitudes(emp.getSolicitudes());

                if (solicitudes.isEmpty()) {
                    PdfPCell sin = new PdfPCell(new Phrase(
                            "Sin solicitudes en el rango seleccionado",
                            ReporteUtil.fuenteTablaData()
                    ));
                    sin.setColspan(9);
                    sin.setHorizontalAlignment(Element.ALIGN_CENTER);
                    sin.setVerticalAlignment(Element.ALIGN_MIDDLE);
                    sin.setPadding(6f);
                    tabla.addCell(sin);
                } else {
                    int i = 1;

                    for (SolicitudVacacionReporteDTO s : solicitudes) {
                        Color fondo = i % 2 == 0 ? zebra : Color.WHITE;

                        tabla.addCell(cellCenter(String.valueOf(long0(s.getSolicitud())), fondo));
                        tabla.addCell(cellCenter(formatearFecha(s.getDesde()), fondo));
                        tabla.addCell(cellCenter(formatearFecha(s.getHasta()), fondo));
                        tabla.addCell(cellCenter(String.valueOf(num0(s.getDias_l_v())), fondo));
                        tabla.addCell(cellCenter(String.valueOf(num0(s.getDias_s_d())), fondo));
                        tabla.addCell(cellCenter(safe(s.getAutorizado()), fondo));
                        tabla.addCell(cellCenter(safe(s.getEstado_texto()), fondo));
                        tabla.addCell(cellLeft(safe(s.getAutoriza()), fondo));
                        tabla.addCell(cellCenter(formatearFechaHora(s.getFecha_autorizacion()), fondo));

                        i++;
                    }
                }

                tabla.setSpacingAfter(10f);
                document.add(tabla);
            }

            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteSolicitudesVacaciones.pdf", e);
        } finally {
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

    // =========================================================================================
    // EXCEL
    // =========================================================================================

    public byte[] generarReporteSolicitudesVacacionesExcel(ReporteSolicitudVacacionRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);
            CellStyle estiloEncabezado = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            CellStyle stCabGray = ConfiguracionExcel.crearEstiloCabeceraGris(libro);
            CellStyle stCabActivo = crearEstiloCabeceraEmpleadoActivo(libro, request.getColorPrincipal());

            XSSFSheet hoja = libro.createSheet("Solicitudes_Vacaciones");

            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            hoja.setColumnWidth(0, 5500);
            hoja.setColumnWidth(1, 4500);
            hoja.setColumnWidth(2, 4500);
            hoja.setColumnWidth(3, 3500);
            hoja.setColumnWidth(4, 3500);
            hoja.setColumnWidth(5, 4000);
            hoja.setColumnWidth(6, 5000);
            hoja.setColumnWidth(7, 9000);
            hoja.setColumnWidth(8, 6500);

            int row = 0;

            for (int fila = 0; fila <= 4; fila++) {
                UtilExcel.combinarCeldas(hoja, fila, fila, 1, 8);
            }

            UtilExcel.establecerTexto(
                    hoja,
                    row++,
                    1,
                    UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    row++,
                    1,
                    "REPORTE - SOLICITUDES DE VACACIONES",
                    estiloTitulo
            );

            UtilExcel.establecerTexto(
                    hoja,
                    row++,
                    1,
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta()),
                    estiloTitulo
            );

            row += 2;

            List<SolicitudVacacionEmpleadoDTO> empleados = agruparPorEmpleado(request.getSolicitudes());

            if (empleados.isEmpty()) {
                mergeSafeNoBorder(hoja, row, row, 0, 8, estiloCentroBorde);
                UtilExcel.establecerTexto(hoja, row, 0, "Sin datos para mostrar", estiloCentroBorde);
                libro.write(baos);
                return baos.toByteArray();
            }

            for (SolicitudVacacionEmpleadoDTO emp : empleados) {
                if (emp == null) continue;

                // =============================================================================
                // Cabecera del empleado
                // Se mantienen los mismos datos, combinaciones y distribución.
                // =============================================================================

                mergeSafeNoBorder(hoja, row, row, 0, 2, stCabActivo);
                mergeSafeNoBorder(hoja, row, row, 3, 5, stCabActivo);
                mergeSafeNoBorder(hoja, row, row, 6, 8, stCabActivo);

                UtilExcel.establecerTexto(hoja, row, 0, "CIUDAD: " + safe(emp.getCiudad()), stCabActivo);
                UtilExcel.establecerTexto(hoja, row, 3, "C.C.: " + safe(emp.getIdentificacion()), stCabActivo);
                UtilExcel.establecerTexto(hoja, row, 6, "COD: " + safe(emp.getCodigo()), stCabActivo);
                row++;

                mergeSafeNoBorder(hoja, row, row, 0, 8, stCabActivo);
                UtilExcel.establecerTexto(hoja, row++, 0, "EMPLEADO: " + safe(emp.getEmpleado()), stCabActivo);

                mergeSafeNoBorder(hoja, row, row, 0, 2, stCabGray);
                mergeSafeNoBorder(hoja, row, row, 3, 5, stCabGray);
                mergeSafeNoBorder(hoja, row, row, 6, 8, stCabGray);

                UtilExcel.establecerTexto(hoja, row, 0, "RÉGIMEN LABORAL: " + safe(emp.getRegimen()), stCabGray);
                UtilExcel.establecerTexto(hoja, row, 3, "DEPARTAMENTO: " + safe(emp.getDepartamento()), stCabGray);
                UtilExcel.establecerTexto(hoja, row, 6, "CARGO: " + safe(emp.getCargo()), stCabGray);
                row++;

                // =============================================================================
                // Encabezado de las solicitudes
                // Mismas columnas, únicamente con el estilo estándar de encabezado.
                // =============================================================================

                Row head = UtilExcel.asegurarFila(hoja, row++);
                UtilExcel.establecerTexto(head, 0, "SOLICITUD", estiloEncabezado);
                UtilExcel.establecerTexto(head, 1, "DESDE", estiloEncabezado);
                UtilExcel.establecerTexto(head, 2, "HASTA", estiloEncabezado);
                UtilExcel.establecerTexto(head, 3, "DÍAS L-V", estiloEncabezado);
                UtilExcel.establecerTexto(head, 4, "DÍAS S-D", estiloEncabezado);
                UtilExcel.establecerTexto(head, 5, "AUTORIZADO", estiloEncabezado);
                UtilExcel.establecerTexto(head, 6, "ESTADO", estiloEncabezado);
                UtilExcel.establecerTexto(head, 7, "AUTORIZA", estiloEncabezado);
                UtilExcel.establecerTexto(head, 8, "FECHA DE AUTORIZACIÓN", estiloEncabezado);
                head.setHeightInPoints(22f);

                List<SolicitudVacacionReporteDTO> solicitudes = ordenarSolicitudes(emp.getSolicitudes());

                if (solicitudes.isEmpty()) {
                    mergeSafe(hoja, row, row, 0, 8, estiloCentroBorde);
                    UtilExcel.establecerTexto(
                            hoja,
                            row++,
                            0,
                            "Sin solicitudes en el rango seleccionado",
                            estiloCentroBorde
                    );
                } else {
                    for (SolicitudVacacionReporteDTO s : solicitudes) {
                        Row rr = UtilExcel.asegurarFila(hoja, row++);

                        UtilExcel.establecerValor(rr, 0, long0(s.getSolicitud()), estiloCentroBorde);
                        UtilExcel.establecerTexto(rr, 1, formatearFecha(s.getDesde()), estiloCentroBorde);
                        UtilExcel.establecerTexto(rr, 2, formatearFecha(s.getHasta()), estiloCentroBorde);
                        UtilExcel.establecerValor(rr, 3, num0(s.getDias_l_v()), estiloCentroBorde);
                        UtilExcel.establecerValor(rr, 4, num0(s.getDias_s_d()), estiloCentroBorde);
                        UtilExcel.establecerTexto(rr, 5, safe(s.getAutorizado()), estiloCentroBorde);
                        UtilExcel.establecerTexto(rr, 6, safe(s.getEstado_texto()), estiloCentroBorde);
                        UtilExcel.establecerTexto(rr, 7, safe(s.getAutoriza()), estiloIzqBorde);
                        UtilExcel.establecerTexto(
                                rr,
                                8,
                                formatearFechaHora(s.getFecha_autorizacion()),
                                estiloCentroBorde
                        );
                    }
                }

                row += 2;
            }

            hoja.createFreezePane(0, 5);

            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteSolicitudesVacaciones.xlsx", e);
        }
    }

    // =========================================================================================
    // Helpers agrupación
    // =========================================================================================

    private List<SolicitudVacacionEmpleadoDTO> agruparPorEmpleado(List<SolicitudVacacionReporteDTO> filas) {
        if (filas == null || filas.isEmpty()) return new ArrayList<>();

        Map<Long, SolicitudVacacionEmpleadoDTO> map = new LinkedHashMap<>();

        for (SolicitudVacacionReporteDTO row : filas) {
            if (row == null) continue;

            Long id = row.getId_empleado() == null ? 0L : row.getId_empleado();

            if (!map.containsKey(id)) {
                SolicitudVacacionEmpleadoDTO emp = new SolicitudVacacionEmpleadoDTO();

                emp.setId_empleado(id);
                emp.setIdentificacion(row.getIdentificacion());
                emp.setCodigo(row.getCodigo());
                emp.setNombre(row.getNombre());
                emp.setApellido(row.getApellido());
                emp.setEmpleado(
                        safe(row.getEmpleado()).isBlank()
                                ? (safe(row.getApellido()) + " " + safe(row.getNombre())).trim()
                                : row.getEmpleado()
                );
                emp.setCiudad(row.getCiudad());
                emp.setSucursal(row.getSucursal());
                emp.setRegimen(row.getRegimen());
                emp.setDepartamento(row.getDepartamento());
                emp.setCargo(row.getCargo());
                emp.setRol(row.getRol());
                emp.setSolicitudes(new ArrayList<>());

                map.put(id, emp);
            }

            map.get(id).getSolicitudes().add(row);
        }

        return map.values().stream()
                .sorted(Comparator.comparing((SolicitudVacacionEmpleadoDTO e) -> safe(e.getApellido()))
                        .thenComparing(e -> safe(e.getNombre())))
                .collect(Collectors.toList());
    }

    private List<SolicitudVacacionReporteDTO> ordenarSolicitudes(List<SolicitudVacacionReporteDTO> solicitudes) {
        if (solicitudes == null) return new ArrayList<>();

        return solicitudes.stream()
                .filter(Objects::nonNull)
                .sorted(Comparator.comparing((SolicitudVacacionReporteDTO s) -> safe(s.getDesde()))
                        .thenComparing(s -> long0(s.getSolicitud())))
                .collect(Collectors.toList());
    }

    // =========================================================================================
    // Helpers PDF
    // =========================================================================================

    private PdfPCell hCell(String text, Color bg) {
        return ReporteUtil.crearCelda(text, ReporteUtil.fuenteEncabezadoTablaData(), bg);
    }

    private PdfPCell cellCenter(String text, Color bg) {
        PdfPCell c = ReporteUtil.crearCelda(safe(text), ReporteUtil.fuenteTablaData(), bg);
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setVerticalAlignment(Element.ALIGN_MIDDLE);
        return c;
    }

    private PdfPCell cellLeft(String text, Color bg) {
        PdfPCell c = ReporteUtil.crearCelda(safe(text), ReporteUtil.fuenteTablaData(), bg);
        c.setHorizontalAlignment(Element.ALIGN_LEFT);
        c.setVerticalAlignment(Element.ALIGN_MIDDLE);
        return c;
    }

    // =========================================================================================
    // Helpers Excel
    // =========================================================================================

    private void mergeSafe(XSSFSheet sh, int r1, int r2, int c1, int c2, CellStyle estilo) {
        if (r1 == r2 && c1 == c2) {
            Row row = UtilExcel.asegurarFila(sh, r1);
            Cell cell = row.getCell(c1);
            if (cell == null) cell = row.createCell(c1);
            if (estilo != null) cell.setCellStyle(estilo);
            return;
        }

        int rr1 = Math.min(r1, r2);
        int rr2 = Math.max(r1, r2);
        int cc1 = Math.min(c1, c2);
        int cc2 = Math.max(c1, c2);

        CellRangeAddress region = new CellRangeAddress(rr1, rr2, cc1, cc2);

        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress existente = sh.getMergedRegion(i);

            if (existente.formatAsString().equals(region.formatAsString())) {
                UtilExcel.aplicarEstiloARegion(sh, rr1, rr2, cc1, cc2, estilo, true);
                aplicarBordeRegion(sh, region);
                return;
            }
        }

        sh.addMergedRegion(region);
        UtilExcel.aplicarEstiloARegion(sh, rr1, rr2, cc1, cc2, estilo, true);
        aplicarBordeRegion(sh, region);
    }

    private void mergeSafeNoBorder(XSSFSheet sh, int r1, int r2, int c1, int c2, CellStyle estilo) {
        if (r1 == r2 && c1 == c2) {
            Row row = UtilExcel.asegurarFila(sh, r1);
            Cell cell = row.getCell(c1);
            if (cell == null) cell = row.createCell(c1);
            if (estilo != null) cell.setCellStyle(estilo);
            return;
        }

        int rr1 = Math.min(r1, r2);
        int rr2 = Math.max(r1, r2);
        int cc1 = Math.min(c1, c2);
        int cc2 = Math.max(c1, c2);

        CellRangeAddress region = new CellRangeAddress(rr1, rr2, cc1, cc2);
        sh.addMergedRegion(region);

        UtilExcel.aplicarEstiloARegion(sh, rr1, rr2, cc1, cc2, estilo, true);
        RegionUtil.setBorderTop(BorderStyle.NONE, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.NONE, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.NONE, region, sh);
        RegionUtil.setBorderRight(BorderStyle.NONE, region, sh);
    }

    private void aplicarBordeRegion(Sheet sh, CellRangeAddress region) {
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);

        short color = IndexedColors.GREY_50_PERCENT.getIndex();

        RegionUtil.setTopBorderColor(color, region, sh);
        RegionUtil.setBottomBorderColor(color, region, sh);
        RegionUtil.setLeftBorderColor(color, region, sh);
        RegionUtil.setRightBorderColor(color, region, sh);
    }

    private CellStyle crearEstiloCabeceraEmpleadoActivo(XSSFWorkbook libro, String colorHex) {
        CellStyle estilo = libro.createCellStyle();
        estilo.cloneStyleFrom(ConfiguracionExcel.crearEstiloCabeceraGris(libro));

        org.apache.poi.xssf.usermodel.XSSFColor color = new org.apache.poi.xssf.usermodel.XSSFColor(
                ReporteUtil.convertirHexAColor(colorHex),
                null
        );

        ((org.apache.poi.xssf.usermodel.XSSFCellStyle) estilo).setFillForegroundColor(color);
        estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return estilo;
    }

    // =========================================================================================
    // Helpers util
    // =========================================================================================

    private String safe(Object v) {
        if (v == null) return "";

        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private long long0(Long v) {
        return v == null ? 0L : v;
    }

    private int num0(Number n) {
        return n == null ? 0 : n.intValue();
    }

    private String formatearFecha(String fecha) {
        String f = safe(fecha);
        if (f.isBlank()) return "";

        try {
            if (f.contains("T")) {
                return OffsetDateTime.parse(f).format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
            }
        } catch (Exception ignore) {
        }

        try {
            return LocalDate.parse(f).format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (Exception ignore) {
        }

        try {
            return LocalDateTime.parse(f).format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (Exception ignore) {
        }

        return f;
    }

    private String formatearFechaHora(String fecha) {
        String f = safe(fecha);
        if (f.isBlank()) return "";

        try {
            if (f.contains("T")) {
                return OffsetDateTime.parse(f).format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
            }
        } catch (Exception ignore) {
        }

        try {
            return LocalDateTime.parse(f).format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
        } catch (Exception ignore) {
        }

        try {
            return LocalDate.parse(f).format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        } catch (Exception ignore) {
        }

        return f;
    }
}