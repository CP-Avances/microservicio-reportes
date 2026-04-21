package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.reporteSolicitudesPermisos.ReporteSolicitudesPermisosRequest;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesPermisos.SolicitudPermisoReporteDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openpdf.text.*;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class ReporteConsolidadoSolicitudPermisoService {

    // =========================================================================================
    // PDF
    // =========================================================================================
    public byte[] generarReporteSolicitudesPermisosPDF(ReporteSolicitudesPermisosRequest request) {
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
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(safe(request.getEmpresa())));
            document.add(ReporteUtil.crearTituloReporte("REPORTE DE SOLICITUDES DE PERMISOS"));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta())
            ));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color zebra = ReporteUtil.colorZebraClaro();

            List<SolicitudPermisoReporteDTO> filas = request.getSolicitudes();
            if (filas == null || filas.isEmpty()) {
                Paragraph p = new Paragraph("Sin datos para mostrar", ReporteUtil.fuenteTexto());
                p.setAlignment(Element.ALIGN_CENTER);
                p.setSpacingBefore(20f);
                document.add(p);
                document.close();
                return baos.toByteArray();
            }

            List<GrupoEmpleado> grupos = agruparPorEmpleado(filas);

            final float[] W_EMP_INFO = {4f, 4f, 2f};
            final float[] W_TABLA = {1.5f, 2.0f, 3.2f, 1.8f, 1.8f, 1.0f, 1.0f, 1.7f, 1.2f, 2.5f, 2.2f};

            for (GrupoEmpleado emp : grupos) {
                PdfPTable infoEmp = new PdfPTable(3);
                infoEmp.setWidthPercentage(100);
                infoEmp.setWidths(W_EMP_INFO);

                infoEmp.addCell(celdaInfoMixtaLocal("CIUDAD:", emp.ciudad, colorPrincipal));
                infoEmp.addCell(celdaInfoMixtaLocal("C.C.:", emp.identificacion, colorPrincipal));
                infoEmp.addCell(celdaInfoMixtaLocal("COD:", emp.codigo, colorPrincipal));

                PdfPCell cNombre = celdaInfoMixtaLocal("EMPLEADO:", emp.empleado, colorPrincipal);
                cNombre.setColspan(3);
                infoEmp.addCell(cNombre);

                PdfPCell cReg = celdaInfoMixtaLocal("RÉGIMEN LABORAL:", emp.regimen, Color.WHITE);
                cReg.setColspan(3);
                infoEmp.addCell(cReg);

                PdfPCell cDep = celdaInfoMixtaLocal("DEPARTAMENTO:", emp.departamento, Color.WHITE);
                cDep.setColspan(3);
                infoEmp.addCell(cDep);

                PdfPCell cCargo = celdaInfoMixtaLocal("CARGO:", emp.cargo, Color.WHITE);
                cCargo.setColspan(3);
                infoEmp.addCell(cCargo);

                PdfPTable contInfoEmp = new PdfPTable(1);
                contInfoEmp.setWidthPercentage(100);
                PdfPCell wrap = new PdfPCell(infoEmp);
                wrap.setBorder(Rectangle.BOX);
                wrap.setPadding(0);
                contInfoEmp.addCell(wrap);
                contInfoEmp.setSpacingBefore(4f);
                contInfoEmp.setSpacingAfter(8f);

                document.add(contInfoEmp);

                PdfPTable tabla = new PdfPTable(11);
                tabla.setWidthPercentage(100);
                tabla.setWidths(W_TABLA);

                tabla.addCell(hCell("Solicitud", colorPrincipal));
                tabla.addCell(hCell("Fecha Solicitud", colorPrincipal));
                tabla.addCell(hCell("Tipo Permiso", colorPrincipal));
                tabla.addCell(hCell("Desde", colorPrincipal));
                tabla.addCell(hCell("Hasta", colorPrincipal));
                tabla.addCell(hCell("Días", colorPrincipal));
                tabla.addCell(hCell("Horas", colorPrincipal));
                tabla.addCell(hCell("Estado", colorPrincipal));
                tabla.addCell(hCell("Autorizado", colorPrincipal));
                tabla.addCell(hCell("Autoriza", colorPrincipal));
                tabla.addCell(hCell("Fecha Autorización", colorPrincipal));

                int i = 1;
                for (SolicitudPermisoReporteDTO s : emp.solicitudes) {
                    Color fondo = (i % 2 == 0) ? zebra : Color.WHITE;

                    tabla.addCell(cellCenter(String.valueOf(long0(s.getCodigo_solicitud())), fondo));
                    tabla.addCell(cellCenter(formatFechaHora(s.getFecha_solicitud()), fondo));
                    tabla.addCell(cellLeft(safe(s.getTipo_permiso()), fondo));
                    tabla.addCell(cellCenter(formatFecha(s.getFecha_inicio()), fondo));
                    tabla.addCell(cellCenter(formatFecha(s.getFecha_fin()), fondo));
                    tabla.addCell(cellCenter(fmtDec2(s.getDias()), fondo));
                    tabla.addCell(cellCenter(fmtDec2(s.getHoras()), fondo));
                    tabla.addCell(cellCenter(safe(s.getEstado_texto()), fondo));
                    tabla.addCell(cellCenter(safe(s.getAutorizado()), fondo));
                    tabla.addCell(cellLeft(safe(s.getAutoriza()), fondo));
                    tabla.addCell(cellCenter(formatFechaHora(s.getFecha_autorizacion()), fondo));

                    i++;
                }

                tabla.setSpacingAfter(14f);
                document.add(tabla);
            }

            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteSolicitudesPermisos.pdf", e);
        } finally {
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

    // =========================================================================================
    // EXCEL
    // =========================================================================================
    public byte[] generarReporteSolicitudesPermisosExcel(ReporteSolicitudesPermisosRequest request) {
        try (XSSFWorkbook libro = new XSSFWorkbook();
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);
            CellStyle stCabGray = ConfiguracionExcel.crearEstiloCabeceraGris(libro);
            CellStyle stCabActivo = crearEstiloCabeceraEmpleadoActivo(libro, request.getColorPrincipal());
            CellStyle stHeadBlue = ConfiguracionExcel.crearEstiloHeaderAzul(libro);

            XSSFSheet sh = libro.createSheet("Solicitudes_Permisos");

            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                insertarLogoA1A4(libro, sh, logo);
            }

            sh.setColumnWidth(0, 5000);  // Solicitud
            sh.setColumnWidth(1, 5500);  // Fecha Solicitud
            sh.setColumnWidth(2, 10000); // Tipo permiso
            sh.setColumnWidth(3, 4500);  // Desde
            sh.setColumnWidth(4, 4500);  // Hasta
            sh.setColumnWidth(5, 3000);  // Días
            sh.setColumnWidth(6, 3000);  // Horas
            sh.setColumnWidth(7, 5000);  // Estado
            sh.setColumnWidth(8, 4000);  // Autorizado
            sh.setColumnWidth(9, 10000); // Autoriza
            sh.setColumnWidth(10, 6500); // Fecha autorización

            int row = 0;
            int C0 = 0;
            int C_LAST = 10;

            mergeSafeNoBorder(sh, row, row, C0, C_LAST, estiloTitulo);
            UtilExcel.establecerTexto(sh, row++, C0, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())), estiloTitulo);

            mergeSafeNoBorder(sh, row, row, C0, C_LAST, estiloTitulo);
            UtilExcel.establecerTexto(sh, row++, C0, "REPORTE DE SOLICITUDES DE PERMISOS", estiloTitulo);

            mergeSafeNoBorder(sh, row, row, C0, C_LAST, estiloTitulo);
            UtilExcel.establecerTexto(
                    sh,
                    row++,
                    C0,
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta()),
                    estiloTitulo
            );

            row++;

            List<SolicitudPermisoReporteDTO> filas = request.getSolicitudes();
            if (filas == null) {
                filas = new ArrayList<>();
            }

            List<GrupoEmpleado> grupos = agruparPorEmpleado(filas);

            for (GrupoEmpleado emp : grupos) {
                int rCabIni = row;

                mergeSafeNoBorder(sh, row, row, 0, 3, stCabActivo);
                mergeSafeNoBorder(sh, row, row, 4, 7, stCabActivo);
                mergeSafeNoBorder(sh, row, row, 8, 10, stCabActivo);

                UtilExcel.establecerTexto(sh, row, 0, "CIUDAD: " + safe(emp.ciudad), stCabActivo);
                UtilExcel.establecerTexto(sh, row, 4, "C.C.: " + safe(emp.identificacion), stCabActivo);
                UtilExcel.establecerTexto(sh, row, 8, "COD: " + safe(emp.codigo), stCabActivo);
                row++;

                mergeSafeNoBorder(sh, row, row, 0, 10, stCabActivo);
                UtilExcel.establecerTexto(sh, row, 0, "EMPLEADO: " + safe(emp.empleado), stCabActivo);
                row++;

                mergeSafeNoBorder(sh, row, row, 0, 10, stCabGray);
                UtilExcel.establecerTexto(sh, row, 0, "RÉGIMEN LABORAL: " + safe(emp.regimen), stCabGray);
                row++;

                mergeSafeNoBorder(sh, row, row, 0, 10, stCabGray);
                UtilExcel.establecerTexto(sh, row, 0, "DEPARTAMENTO: " + safe(emp.departamento), stCabGray);
                row++;

                mergeSafeNoBorder(sh, row, row, 0, 10, stCabGray);
                UtilExcel.establecerTexto(sh, row, 0, "CARGO: " + safe(emp.cargo), stCabGray);
                row++;

                int rCabFin = row - 1;
                limpiarBordesEnRegion(sh, rCabIni, rCabFin, 0, 10);
                bordeExternoGrueso(sh, new CellRangeAddress(rCabIni, rCabFin, 0, 10));

                Row h = UtilExcel.asegurarFila(sh, row++);
                UtilExcel.establecerTexto(h, 0, "Solicitud", stHeadBlue);
                UtilExcel.establecerTexto(h, 1, "Fecha Solicitud", stHeadBlue);
                UtilExcel.establecerTexto(h, 2, "Tipo Permiso", stHeadBlue);
                UtilExcel.establecerTexto(h, 3, "Desde", stHeadBlue);
                UtilExcel.establecerTexto(h, 4, "Hasta", stHeadBlue);
                UtilExcel.establecerTexto(h, 5, "Días", stHeadBlue);
                UtilExcel.establecerTexto(h, 6, "Horas", stHeadBlue);
                UtilExcel.establecerTexto(h, 7, "Estado", stHeadBlue);
                UtilExcel.establecerTexto(h, 8, "Autorizado", stHeadBlue);
                UtilExcel.establecerTexto(h, 9, "Autoriza", stHeadBlue);
                UtilExcel.establecerTexto(h, 10, "Fecha Autorización", stHeadBlue);

                for (SolicitudPermisoReporteDTO s : emp.solicitudes) {
                    Row rr = UtilExcel.asegurarFila(sh, row++);

                    UtilExcel.establecerValor(rr, 0, long0(s.getCodigo_solicitud()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 1, formatFechaHora(s.getFecha_solicitud()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 2, safe(s.getTipo_permiso()), estiloIzqBorde);
                    UtilExcel.establecerTexto(rr, 3, formatFecha(s.getFecha_inicio()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 4, formatFecha(s.getFecha_fin()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 5, fmtDec2(s.getDias()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 6, fmtDec2(s.getHoras()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 7, safe(s.getEstado_texto()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 8, safe(s.getAutorizado()), estiloCentroBorde);
                    UtilExcel.establecerTexto(rr, 9, safe(s.getAutoriza()), estiloIzqBorde);
                    UtilExcel.establecerTexto(rr, 10, formatFechaHora(s.getFecha_autorizacion()), estiloCentroBorde);
                }

                row += 2;
            }

            sh.createFreezePane(0, 4);

            libro.write(baos);
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteSolicitudesPermisos.xlsx", e);
        }
    }

    // =========================================================================================
    // HELPERS DE AGRUPACIÓN
    // =========================================================================================
    private List<GrupoEmpleado> agruparPorEmpleado(List<SolicitudPermisoReporteDTO> filas) {
        Map<Long, GrupoEmpleado> map = new LinkedHashMap<>();

        List<SolicitudPermisoReporteDTO> ordenadas = filas.stream()
                .filter(Objects::nonNull)
                .sorted(Comparator
                        .comparing((SolicitudPermisoReporteDTO x) -> safe(x.getApellido()))
                        .thenComparing(x -> safe(x.getNombre()))
                        .thenComparing(x -> safe(x.getFecha_inicio()))
                        .thenComparing(x -> long0(x.getSolicitud())))
                .collect(Collectors.toList());

        for (SolicitudPermisoReporteDTO row : ordenadas) {
            Long id = row.getId_empleado() != null ? row.getId_empleado() : 0L;

            if (!map.containsKey(id)) {
                GrupoEmpleado g = new GrupoEmpleado();
                g.idEmpleado = id;
                g.identificacion = safe(row.getIdentificacion());
                g.codigo = safe(row.getCodigo());
                g.empleado = !safe(row.getEmpleado()).isBlank()
                        ? safe(row.getEmpleado())
                        : (safe(row.getApellido()) + " " + safe(row.getNombre())).trim();
                g.ciudad = safe(row.getCiudad());
                g.regimen = safe(row.getRegimen());
                g.departamento = safe(row.getDepartamento());
                g.cargo = safe(row.getCargo());
                g.solicitudes = new ArrayList<>();
                map.put(id, g);
            }

            map.get(id).solicitudes.add(row);
        }

        return new ArrayList<>(map.values());
    }

    private static class GrupoEmpleado {
        Long idEmpleado;
        String identificacion;
        String codigo;
        String empleado;
        String ciudad;
        String regimen;
        String departamento;
        String cargo;
        List<SolicitudPermisoReporteDTO> solicitudes;
    }

    // =========================================================================================
    // HELPERS PDF
    // =========================================================================================
    private PdfPCell hCell(String text, Color bg) {
        return ReporteUtil.crearCelda(text, ReporteUtil.fuenteEncabezado(), bg);
    }

    private PdfPCell cellCenter(String text, Color bg) {
        PdfPCell c = new PdfPCell(new Phrase(safe(text), ReporteUtil.fuenteTablaData()));
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setVerticalAlignment(Element.ALIGN_MIDDLE);
        c.setBackgroundColor(bg);
        c.setPadding(4f);
        c.setBorder(Rectangle.BOX);
        c.setBorderColor(new Color(80, 80, 80));
        c.setBorderWidth(0.6f);
        return c;
    }

    private PdfPCell cellLeft(String text, Color bg) {
        PdfPCell c = new PdfPCell(new Phrase(safe(text), ReporteUtil.fuenteTablaData()));
        c.setHorizontalAlignment(Element.ALIGN_LEFT);
        c.setVerticalAlignment(Element.ALIGN_MIDDLE);
        c.setBackgroundColor(bg);
        c.setPadding(4f);
        c.setBorder(Rectangle.BOX);
        c.setBorderColor(new Color(80, 80, 80));
        c.setBorderWidth(0.6f);
        return c;
    }

    private PdfPCell celdaInfoMixtaLocal(String etiqueta, String valor, Color fondo) {
        Phrase contenido = new Phrase();
        contenido.add(new Chunk(safe(etiqueta) + " ", ReporteUtil.fuenteEncabezadoTablaData()));
        contenido.add(new Chunk(safe(valor), ReporteUtil.fuenteTablaData()));

        PdfPCell celda = new PdfPCell(contenido);
        celda.setBackgroundColor(fondo);
        celda.setPadding(2f);
        celda.setBorder(Rectangle.NO_BORDER);
        return celda;
    }

    // =========================================================================================
    // HELPERS EXCEL
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
            CellRangeAddress ex = sh.getMergedRegion(i);
            if (ex.formatAsString().equals(region.formatAsString())) {
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

        short col = IndexedColors.GREY_50_PERCENT.getIndex();
        RegionUtil.setTopBorderColor(col, region, sh);
        RegionUtil.setBottomBorderColor(col, region, sh);
        RegionUtil.setLeftBorderColor(col, region, sh);
        RegionUtil.setRightBorderColor(col, region, sh);
    }

    private void limpiarBordesEnRegion(XSSFSheet sh, int r1, int r2, int c1, int c2) {
        int rr1 = Math.min(r1, r2), rr2 = Math.max(r1, r2);
        int cc1 = Math.min(c1, c2), cc2 = Math.max(c1, c2);

        for (int r = rr1; r <= rr2; r++) {
            Row row = UtilExcel.asegurarFila(sh, r);
            for (int c = cc1; c <= cc2; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);

                CellStyle st = cell.getCellStyle();
                if (st == null) continue;

                CellStyle nuevo = sh.getWorkbook().createCellStyle();
                nuevo.cloneStyleFrom(st);
                nuevo.setBorderTop(BorderStyle.NONE);
                nuevo.setBorderBottom(BorderStyle.NONE);
                nuevo.setBorderLeft(BorderStyle.NONE);
                nuevo.setBorderRight(BorderStyle.NONE);

                cell.setCellStyle(nuevo);
            }
        }
    }

    private void bordeExternoGrueso(Sheet sh, CellRangeAddress region) {
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);
    }

    private void insertarLogoA1A4(Workbook wb, Sheet hoja, byte[] imagenBytes) {
        if (imagenBytes == null || imagenBytes.length == 0) return;

        int idx = wb.addPicture(imagenBytes, Workbook.PICTURE_TYPE_PNG);
        Drawing<?> drawing = hoja.createDrawingPatriarch();
        CreationHelper helper = wb.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();

        anchor.setCol1(0);
        anchor.setCol2(1);
        anchor.setRow1(0);
        anchor.setRow2(4);
        anchor.setDx1(0);
        anchor.setDx2(0);
        anchor.setDy1(0);
        anchor.setDy2(0);
        anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

        drawing.createPicture(anchor, idx);
    }

    private CellStyle crearEstiloCabeceraEmpleadoActivo(XSSFWorkbook libro, String colorHex) {
        CellStyle estilo = libro.createCellStyle();
        estilo.cloneStyleFrom(ConfiguracionExcel.crearEstiloCabeceraGris(libro));

        org.apache.poi.xssf.usermodel.XSSFColor color =
                new org.apache.poi.xssf.usermodel.XSSFColor(
                        ReporteUtil.convertirHexAColor(colorHex),
                        null
                );

        ((org.apache.poi.xssf.usermodel.XSSFCellStyle) estilo).setFillForegroundColor(color);
        estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return estilo;
    }

    // =========================================================================================
    // HELPERS GENERALES
    // =========================================================================================
    private String safe(Object v) {
        if (v == null) return "";
        String s = String.valueOf(v).trim();
        return "null".equalsIgnoreCase(s) ? "" : s;
    }

    private long long0(Long v) {
        return v == null ? 0L : v;
    }

    private String fmtDec2(Double n) {
        if (n == null) return "0,00";
        return String.format(Locale.US, "%.2f", n).replace('.', ',');
    }

    private String formatFecha(String fecha) {
        String v = safe(fecha);
        if (v.isBlank()) return "";
        if (v.length() >= 10 && v.matches("^\\d{4}-\\d{2}-\\d{2}.*")) {
            return v.substring(8, 10) + "/" + v.substring(5, 7) + "/" + v.substring(0, 4);
        }
        return v;
    }

    private String formatFechaHora(String fecha) {
        String v = safe(fecha);
        if (v.isBlank()) return "";

        try {
            if (v.matches("^\\d{4}-\\d{2}-\\d{2}$")) {
                return v.substring(8, 10) + "/" + v.substring(5, 7) + "/" + v.substring(0, 4);
            }

            if (v.length() >= 16 && v.matches("^\\d{4}-\\d{2}-\\d{2}.*")) {
                String f = v.substring(8, 10) + "/" + v.substring(5, 7) + "/" + v.substring(0, 4);
                String h = v.substring(11, 16);
                return f + " " + h;
            }
        } catch (Exception ignore) {
        }

        return v;
    }
}