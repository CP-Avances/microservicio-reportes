package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.ReporteSolicitudVacacionRequest;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.SolicitudVacacionAprobacionDTO;
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
import org.openpdf.text.DocumentException;
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
import java.util.stream.Collectors;

@Service
public class ReporteConsolidadoSolicitudVacacionService {

    // =========================================================================================
    // PDF
    // =========================================================================================

    public byte[] generarReporteSolicitudesVacacionesPDF(
            ReporteSolicitudVacacionRequest request) {

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {

            if (request == null) {
                throw new IllegalArgumentException("request nulo");
            }

            baos = new ByteArrayOutputStream();

            document = new Document(
                    PageSize.A4.rotate(),
                    30,
                    30,
                    30,
                    50);

            writer = PdfWriter.getInstance(
                    document,
                    baos);

            writer.setPageEvent(
                    new ConfiguracionPaginaPDF(
                            safe(request.getUsuario()),
                            safe(request.getFraseMarcaAgua()),
                            safe(request.getColorPrincipal())));

            document.open();

            // ============================================================
            // LOGO
            // ============================================================

            Image logo = ReporteUtil.obtenerLogo(
                    request.getLogoBase64());

            if (logo != null) {
                document.add(logo);
            }

            // ============================================================
            // TÍTULOS
            // ============================================================

            document.add(
                    ReporteUtil.crearTituloEmpresa(
                            safe(request.getEmpresa())));

            document.add(
                    ReporteUtil.crearTituloReporte(
                            "REPORTE DE SOLICITUDES DE VACACIONES"));

            document.add(
                    ReporteUtil.crearTituloPeriodo(
                            "PERIODO DEL: "
                                    + safe(request.getFechaDesde())
                                    + " AL "
                                    + safe(request.getFechaHasta())));

            // ============================================================
            // COLORES
            // ============================================================

            Color colorPrincipal = ReporteUtil.convertirHexAColor(
                    request.getColorPrincipal());

            Color colorSecundario = ReporteUtil.convertirHexAColor(
                    request.getColorSecundario());

            Color zebraColor = ReporteUtil.colorZebraClaro();

            // ============================================================
            // AGRUPACIÓN
            // ============================================================

            List<SolicitudVacacionEmpleadoDTO> empleados = agruparPorEmpleado(
                    request.getSolicitudes());

            if (empleados == null || empleados.isEmpty()) {

                Paragraph p = new Paragraph(
                        "Sin datos para mostrar",
                        ReporteUtil.fuenteTexto());

                p.setAlignment(
                        Element.ALIGN_CENTER);

                p.setSpacingBefore(20f);

                document.add(p);

                document.close();

                return baos.toByteArray();
            }

            // ============================================================
            // TOTAL DE SOLICITUDES
            // ============================================================

            int totalSolicitudes = empleados.stream()
                    .filter(Objects::nonNull)
                    .mapToInt(
                            e -> e.getSolicitudes() == null
                                    ? 0
                                    : e.getSolicitudes().size())
                    .sum();

            // ============================================================
            // FRANJA LISTA DE SOLICITUDES
            // ============================================================

            PdfPTable tituloTabla = new PdfPTable(2);

            tituloTabla.setWidthPercentage(100);

            tituloTabla.setWidths(
                    new float[] {
                            8f,
                            2f
                    });

            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(
                    new Phrase(
                            "LISTA DE SOLICITUDES",
                            ReporteUtil.fuenteEncabezadoTablaData()));

            celdaTitulo.setBackgroundColor(
                    colorSecundario);

            celdaTitulo.setPadding(5f);

            celdaTitulo.setBorder(
                    Rectangle.TOP
                            | Rectangle.BOTTOM
                            | Rectangle.LEFT);

            tituloTabla.addCell(
                    celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(
                    new Phrase(
                            "N° Registros: "
                                    + totalSolicitudes,
                            ReporteUtil.fuenteEncabezadoTablaData()));

            celdaContador.setBackgroundColor(
                    colorSecundario);

            celdaContador.setHorizontalAlignment(
                    Element.ALIGN_RIGHT);

            celdaContador.setVerticalAlignment(
                    Element.ALIGN_MIDDLE);

            celdaContador.setPadding(5f);

            celdaContador.setBorder(
                    Rectangle.TOP
                            | Rectangle.BOTTOM
                            | Rectangle.RIGHT);

            tituloTabla.addCell(
                    celdaContador);

            document.add(
                    tituloTabla);

            // ============================================================
            // EMPLEADOS
            // ============================================================

            int contadorSolicitud = 0;

            for (SolicitudVacacionEmpleadoDTO emp : empleados) {

                if (emp == null) {
                    continue;
                }

                // ========================================================
                // INFORMACIÓN DEL EMPLEADO
                // MISMO DISEÑO QUE PERMISOS
                // ========================================================

                PdfPTable infoEmpleado = new PdfPTable(3);

                infoEmpleado.setWidthPercentage(100);

                infoEmpleado.setWidths(
                        new float[] {
                                4f,
                                4f,
                                4f
                        });

                infoEmpleado.addCell(
                        ReporteUtil.celdaInfoMixta(
                                "C.C.:",
                                safe(emp.getIdentificacion()),
                                zebraColor));

                infoEmpleado.addCell(
                        ReporteUtil.celdaInfoMixta(
                                "EMPLEADO:",
                                safe(emp.getEmpleado()),
                                zebraColor));

                infoEmpleado.addCell(
                        ReporteUtil.celdaInfoMixta(
                                "COD:",
                                safe(emp.getCodigo()),
                                zebraColor));

                infoEmpleado.addCell(
                        ReporteUtil.celdaInfoMixta(
                                "CIUDAD:",
                                safe(emp.getCiudad()),
                                zebraColor));

                infoEmpleado.addCell(
                        ReporteUtil.celdaInfoMixta(
                                "RÉGIMEN LABORAL:",
                                safe(emp.getRegimen()),
                                zebraColor));

                infoEmpleado.addCell(
                        ReporteUtil.celdaInfoMixta(
                                "DEPARTAMENTO:",
                                safe(emp.getDepartamento()),
                                zebraColor));

                PdfPCell cargo = ReporteUtil.celdaInfoMixta(
                        "CARGO:",
                        safe(emp.getCargo()),
                        zebraColor);

                cargo.setColspan(3);

                infoEmpleado.addCell(
                        cargo);

                PdfPTable contenedora = new PdfPTable(1);

                contenedora.setWidthPercentage(100);

                contenedora.setSpacingAfter(7f);

                PdfPCell contenedor = new PdfPCell(
                        infoEmpleado);

                contenedor.setPadding(0f);

                contenedor.setBorder(
                        Rectangle.BOX);

                contenedora.addCell(
                        contenedor);

                document.add(
                        contenedora);

                // ========================================================
                // SOLICITUDES DEL EMPLEADO
                // ========================================================

                List<SolicitudVacacionReporteDTO> solicitudes = ordenarSolicitudes(
                        emp.getSolicitudes());

                if (solicitudes == null || solicitudes.isEmpty()) {

                    Paragraph p = new Paragraph(
                            "Sin solicitudes registradas hasta la fecha de corte",
                            ReporteUtil.fuenteTexto());

                    p.setAlignment(
                            Element.ALIGN_CENTER);

                    p.setSpacingAfter(10f);

                    document.add(p);

                    continue;
                }

                // ========================================================
                // UNA TABLA POR SOLICITUD
                // ========================================================

                for (SolicitudVacacionReporteDTO solicitud : solicitudes) {

                    if (solicitud == null) {
                        continue;
                    }

                    Color fondo = contadorSolicitud++ % 2 == 0
                            ? Color.WHITE
                            : zebraColor;

                    PdfPTable bloque = new PdfPTable(1);

                    bloque.setWidthPercentage(100);

                    bloque.setSpacingAfter(10f);

                    bloque.setKeepTogether(true);

                    PdfPCell contenido = new PdfPCell();

                    contenido.setBorder(
                            Rectangle.NO_BORDER);

                    contenido.setPadding(0f);

                    // SOLICITUD
                    contenido.addElement(
                            crearTablaSolicitudVacacionPDF(
                                    solicitud,
                                    colorPrincipal,
                                    fondo));

                    // HISTORIAL
                    // SOLO SI EXISTE
                    if (solicitud.getAprobaciones() != null
                            && !solicitud
                                    .getAprobaciones()
                                    .isEmpty()) {

                        contenido.addElement(
                                crearTablaHistorialVacacionPDF(
                                        solicitud,
                                        colorSecundario));
                    }

                    bloque.addCell(
                            contenido);

                    document.add(
                            bloque);
                }
            }

            document.close();

            return baos.toByteArray();

        } catch (IllegalArgumentException e) {

            throw e;

        } catch (Exception e) {

            throw new ReportBuildException(
                    "No se pudo generar ReporteSolicitudesVacaciones.pdf",
                    e);

        } finally {

            if (document != null
                    && document.isOpen()) {

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
            if (logo != null && logo.length > 0)
                UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            hoja.setColumnWidth(0, 5500); // empleado / solicitud
            hoja.setColumnWidth(1, 4500);
            hoja.setColumnWidth(2, 4500);
            hoja.setColumnWidth(3, 3500);
            hoja.setColumnWidth(4, 3500);
            hoja.setColumnWidth(5, 4000);
            hoja.setColumnWidth(6, 5000);
            hoja.setColumnWidth(7, 9000);
            hoja.setColumnWidth(8, 6500);

            int row = 0;

            mergeSafeNoBorder(hoja, row, row, 0, 8, estiloTitulo);
            UtilExcel.establecerTexto(hoja, row++, 0, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);

            UtilExcel.establecerTexto(
                    hoja,
                    row++,
                    1,
                    UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                    estiloTitulo);

            UtilExcel.establecerTexto(
                    hoja,
                    row++,
                    1,
                    "REPORTE - SOLICITUDES DE VACACIONES",
                    estiloTitulo);

            UtilExcel.establecerTexto(
                    hoja,
                    row++,
                    1,
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta()),
                    estiloTitulo);

            row += 2;

            List<SolicitudVacacionEmpleadoDTO> empleados = agruparPorEmpleado(request.getSolicitudes());

            if (empleados.isEmpty()) {
                mergeSafeNoBorder(hoja, row, row, 0, 8, estiloCentroBorde);
                UtilExcel.establecerTexto(hoja, row, 0, "Sin datos para mostrar", estiloCentroBorde);
                libro.write(baos);
                return baos.toByteArray();
            }

            for (SolicitudVacacionEmpleadoDTO emp : empleados) {
                if (emp == null)
                    continue;

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
                    UtilExcel.establecerTexto(hoja, row++, 0, "Sin solicitudes en el rango seleccionado",
                            estiloCentroBorde);
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
                        UtilExcel.establecerTexto(rr, 8, formatearFechaHora(s.getFecha_autorizacion()),
                                estiloCentroBorde);
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
        if (filas == null || filas.isEmpty())
            return new ArrayList<>();

        Map<Long, SolicitudVacacionEmpleadoDTO> empleadosMap = new LinkedHashMap<>();

        for (SolicitudVacacionReporteDTO row : filas) {
            if (row == null)
                continue;

            Long idEmpleado = row.getId_empleado() == null
                    ? 0L
                    : row.getId_empleado();

            // =========================================================
            // AGRUPAR POR EMPLEADO
            // =========================================================
            if (!empleadosMap.containsKey(idEmpleado)) {
                SolicitudVacacionEmpleadoDTO emp = new SolicitudVacacionEmpleadoDTO();
                emp.setId_empleado(idEmpleado);
                emp.setIdentificacion(row.getIdentificacion());
                emp.setCodigo(row.getCodigo());
                emp.setNombre(row.getNombre());
                emp.setApellido(row.getApellido());
                emp.setEmpleado(
                        safe(row.getEmpleado()).isBlank()
                                ? (safe(row.getApellido()) + " " + safe(row.getNombre())).trim()
                                : row.getEmpleado());
                emp.setCiudad(row.getCiudad());
                emp.setSucursal(row.getSucursal());
                emp.setRegimen(row.getRegimen());
                emp.setDepartamento(row.getDepartamento());
                emp.setCargo(row.getCargo());
                emp.setRol(row.getRol());

                emp.setSolicitudes(new ArrayList<>());
                empleadosMap.put(idEmpleado, emp);
            }

            SolicitudVacacionEmpleadoDTO empleado = empleadosMap.get(idEmpleado);

            // =========================================================
            // AGRUPAR POR SOLICITUD
            // =========================================================
            SolicitudVacacionReporteDTO solicitud = empleado.getSolicitudes()
                    .stream()
                    .filter(s -> Objects.equals(s.getSolicitud(), row.getSolicitud()))
                    .findFirst()
                    .orElse(null);

            if (solicitud == null) {
                solicitud = row;
                solicitud.setAprobaciones(new ArrayList<>());
                empleado.getSolicitudes().add(solicitud);
            }

            // =========================================================
            // AGREGAR HISTORIAL DE APROBACIÓN
            // =========================================================
            if (row.getId_historial() != null) {

                boolean yaExiste = solicitud.getAprobaciones()
                        .stream()
                        .anyMatch(a -> Objects.equals(
                                a.getId_historial(),
                                row.getId_historial()));

                if (!yaExiste) {
                    SolicitudVacacionAprobacionDTO aprobacion = new SolicitudVacacionAprobacionDTO();

                    aprobacion.setId_historial(row.getId_historial());
                    aprobacion.setOrden_paso(row.getOrden_paso());

                    aprobacion.setId_departamento_destino(
                            row.getId_departamento_destino());

                    aprobacion.setDepartamento_aprobacion(
                            row.getDepartamento_aprobacion());

                    aprobacion.setDepartamento_nombre(
                            row.getDepartamento_nombre());

                    aprobacion.setId_empleado_aprobador(
                            row.getId_empleado_aprobador());

                    aprobacion.setAutoriza(row.getAutoriza());
                    aprobacion.setEmpleado_nombre(row.getEmpleado_nombre());

                    aprobacion.setAccion(row.getAccion());
                    aprobacion.setEstado_flujo(row.getEstado_flujo());

                    aprobacion.setFecha_autorizacion(
                            row.getFecha_autorizacion());

                    aprobacion.setFecha_hora_accion(
                            row.getFecha_hora_accion());

                    aprobacion.setObservacion(row.getObservacion());
                    aprobacion.setTipo_paso(row.getTipo_paso());
                    aprobacion.setObligatorio(row.getObligatorio());
                    aprobacion.setModo_aprobador(row.getModo_aprobador());
                    aprobacion.setCargo_en_momento(row.getCargo_en_momento());

                    aprobacion.setEs_jefe_en_momento(
                            row.getEs_jefe_en_momento());

                    aprobacion.setHistorial_activo(
                            row.getHistorial_activo());

                    solicitud.getAprobaciones().add(aprobacion);
                }
            }
        }

        List<SolicitudVacacionEmpleadoDTO> empleados = new ArrayList<>(empleadosMap.values());

        // =========================================================
        // ORDENAR EMPLEADOS, SOLICITUDES Y APROBACIONES
        // =========================================================
        empleados.sort(
                Comparator
                        .comparing(
                                (SolicitudVacacionEmpleadoDTO e) -> safe(e.getApellido()))
                        .thenComparing(e -> safe(e.getNombre())));

        for (SolicitudVacacionEmpleadoDTO emp : empleados) {

            emp.getSolicitudes().sort(
                    Comparator
                            .comparing(
                                    (SolicitudVacacionReporteDTO s) -> safe(s.getDesde()))
                            .thenComparing(
                                    s -> long0(s.getSolicitud())));

            for (SolicitudVacacionReporteDTO solicitud : emp.getSolicitudes()) {

                if (solicitud.getAprobaciones() == null) {
                    solicitud.setAprobaciones(new ArrayList<>());
                    continue;
                }

                solicitud.getAprobaciones().sort(
                        Comparator
                                .comparingInt(
                                        (SolicitudVacacionAprobacionDTO a) -> a.getOrden_paso() == null
                                                ? 999
                                                : a.getOrden_paso())
                                .thenComparing(
                                        (SolicitudVacacionAprobacionDTO a) -> safe(a.getFecha_autorizacion())));
            }
        }

        return empleados;
    }

    private List<SolicitudVacacionReporteDTO> ordenarSolicitudes(List<SolicitudVacacionReporteDTO> solicitudes) {
        if (solicitudes == null)
            return new ArrayList<>();

        return solicitudes.stream()
                .filter(Objects::nonNull)
                .sorted(Comparator.comparing((SolicitudVacacionReporteDTO s) -> safe(s.getDesde()))
                        .thenComparing(s -> long0(s.getSolicitud())))
                .collect(Collectors.toList());
    }

    // =========================================================================================
    // Helpers PDF
    // =========================================================================================

    private PdfPTable crearTablaSolicitudVacacionPDF(
            SolicitudVacacionReporteDTO solicitud,
            Color colorPrincipal,
            Color fondo) throws DocumentException {

        final float[] widths = {
                1.4f,
                1.8f,
                1.8f,
                1.2f,
                1.2f,
                1.5f,
                2.0f
        };

        PdfPTable tabla = new PdfPTable(7);

        tabla.setWidthPercentage(100);

        tabla.setWidths(
                widths);

        tabla.setSpacingAfter(0f);

        // ============================================================
        // ENCABEZADOS
        // ============================================================

        tabla.addCell(
                hCell(
                        "SOLICITUD",
                        colorPrincipal));

        tabla.addCell(
                hCell(
                        "DESDE",
                        colorPrincipal));

        tabla.addCell(
                hCell(
                        "HASTA",
                        colorPrincipal));

        tabla.addCell(
                hCell(
                        "DÍAS L-V",
                        colorPrincipal));

        tabla.addCell(
                hCell(
                        "DÍAS S-D",
                        colorPrincipal));

        tabla.addCell(
                hCell(
                        "AUTORIZADO",
                        colorPrincipal));

        tabla.addCell(
                hCell(
                        "ESTADO",
                        colorPrincipal));

        // ============================================================
        // DATOS
        // ============================================================

        tabla.addCell(
                cellCenter(
                        String.valueOf(
                                long0(
                                        solicitud.getSolicitud())),
                        fondo));

        tabla.addCell(
                cellCenter(
                        formatearFecha(
                                solicitud.getDesde()),
                        fondo));

        tabla.addCell(
                cellCenter(
                        formatearFecha(
                                solicitud.getHasta()),
                        fondo));

        tabla.addCell(
                cellCenter(
                        String.valueOf(
                                num0(
                                        solicitud.getDias_l_v())),
                        fondo));

        tabla.addCell(
                cellCenter(
                        String.valueOf(
                                num0(
                                        solicitud.getDias_s_d())),
                        fondo));

        tabla.addCell(
                cellCenter(
                        safe(
                                solicitud.getAutorizado()),
                        fondo));

        tabla.addCell(
                cellCenter(
                        safe(
                                solicitud.getEstado_texto()),
                        fondo));

        return tabla;
    }

    private PdfPTable crearTablaHistorialVacacionPDF(
            SolicitudVacacionReporteDTO solicitud,
            Color colorSecundario) throws DocumentException {

        final float[] widths = {
                0.8f,
                2.6f,
                3.2f,
                1.8f,
                2.3f,
                4.3f
        };

        PdfPTable tabla = new PdfPTable(6);

        tabla.setWidthPercentage(100);

        tabla.setWidths(
                widths);

        tabla.setSpacingBefore(0f);

        tabla.setSpacingAfter(0f);

        // ============================================================
        // ENCABEZADOS
        // ============================================================

        tabla.addCell(
                hCell(
                        "PASO",
                        colorSecundario));

        tabla.addCell(
                hCell(
                        "DEPARTAMENTO",
                        colorSecundario));

        tabla.addCell(
                hCell(
                        "AUTORIZA",
                        colorSecundario));

        tabla.addCell(
                hCell(
                        "ESTADO FLUJO",
                        colorSecundario));

        tabla.addCell(
                hCell(
                        "FECHA AUTORIZACIÓN",
                        colorSecundario));

        tabla.addCell(
                hCell(
                        "OBSERVACIÓN",
                        colorSecundario));

        // ============================================================
        // HISTORIAL
        // ============================================================

        int contador = 1;

        Color zebra = ReporteUtil.colorZebraClaro();

        for (SolicitudVacacionAprobacionDTO aprobacion : solicitud.getAprobaciones()) {

            if (aprobacion == null) {
                continue;
            }

            Color fondo = contador++ % 2 == 0
                    ? zebra
                    : Color.WHITE;

            tabla.addCell(
                    cellCenter(
                            aprobacion.getOrden_paso() == null
                                    ? ""
                                    : String.valueOf(
                                            aprobacion.getOrden_paso()),
                            fondo));

            tabla.addCell(
                    cellLeft(
                            safe(
                                    aprobacion
                                            .getDepartamento_aprobacion()),
                            fondo));

            tabla.addCell(
                    cellLeft(
                            safe(
                                    aprobacion.getAutoriza()),
                            fondo));

            tabla.addCell(
                    cellCenter(
                            safe(
                                    aprobacion.getEstado_flujo()),
                            fondo));

            tabla.addCell(
                    cellCenter(
                            formatearFechaHora(
                                    aprobacion
                                            .getFecha_autorizacion()),
                            fondo));

            tabla.addCell(
                    cellLeft(
                            safe(
                                    aprobacion.getObservacion()).isBlank()
                                            ? "—"
                                            : safe(
                                                    aprobacion.getObservacion()),
                            fondo));
        }

        return tabla;
    }

    private PdfPCell hCell(
            String text,
            Color bg) {

        PdfPCell c = new PdfPCell(
                new Phrase(
                        safe(text),
                        ReporteUtil.fuenteEncabezadoTablaData()));

        c.setHorizontalAlignment(
                Element.ALIGN_CENTER);

        c.setVerticalAlignment(
                Element.ALIGN_MIDDLE);

        c.setBackgroundColor(bg);

        c.setPadding(4f);

        c.setBorder(
                Rectangle.BOX);

        c.setBorderColor(
                new Color(
                        80,
                        80,
                        80));

        c.setBorderWidth(0.8f);

        return c;
    }

    private PdfPCell cellCenter(
            String text,
            Color bg) {

        PdfPCell c = new PdfPCell(
                new Phrase(
                        safe(text),
                        ReporteUtil.fuenteTablaData()));

        c.setHorizontalAlignment(
                Element.ALIGN_CENTER);

        c.setVerticalAlignment(
                Element.ALIGN_MIDDLE);

        c.setBackgroundColor(bg);

        c.setPadding(4f);

        c.setBorder(
                Rectangle.BOX);

        c.setBorderColor(
                new Color(
                        80,
                        80,
                        80));

        c.setBorderWidth(0.6f);

        return c;
    }

    private PdfPCell cellLeft(
            String text,
            Color bg) {

        PdfPCell c = new PdfPCell(
                new Phrase(
                        safe(text),
                        ReporteUtil.fuenteTablaData()));

        c.setHorizontalAlignment(
                Element.ALIGN_LEFT);

        c.setVerticalAlignment(
                Element.ALIGN_MIDDLE);

        c.setBackgroundColor(bg);

        c.setPadding(4f);

        c.setBorder(
                Rectangle.BOX);

        c.setBorderColor(
                new Color(
                        80,
                        80,
                        80));

        c.setBorderWidth(0.6f);

        return c;
    }

    // =========================================================================================
    // Helpers Excel
    // =========================================================================================
    private void mergeSafe(XSSFSheet sh, int r1, int r2, int c1, int c2, CellStyle estilo) {
        if (r1 == r2 && c1 == c2) {
            Row row = UtilExcel.asegurarFila(sh, r1);
            Cell cell = row.getCell(c1);
            if (cell == null)
                cell = row.createCell(c1);
            if (estilo != null)
                cell.setCellStyle(estilo);
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
            if (cell == null)
                cell = row.createCell(c1);
            if (estilo != null)
                cell.setCellStyle(estilo);
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
                null);

        ((org.apache.poi.xssf.usermodel.XSSFCellStyle) estilo).setFillForegroundColor(color);
        estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return estilo;
    }

    // =========================================================================================
    // Helpers util
    // =========================================================================================

    private String safe(Object v) {
        if (v == null)
            return "";
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
        if (f.isBlank())
            return "";

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
        if (f.isBlank())
            return "";

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