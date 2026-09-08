package com.casapazmino.microservicio_reportes.service;

import org.springframework.stereotype.Service;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.ReporteSolicitudVacacionRequest;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.SolicitudVacacionAprobacionDTO;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.SolicitudVacacionEmpleadoDTO;
import com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones.SolicitudVacacionReporteDTO;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import org.apache.poi.ss.usermodel.Row;

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
        final float[] W_EMP_INFO = { 4f, 4f, 2f };
        final float[] W_EMP_INFO_2 = { 4f, 4f, 3f };
        final float[] W_SOL = {
                1.4f, // Solicitud
                1.8f, // Desde
                1.8f, // Hasta
                1.2f, // Días L-V
                1.2f, // Días S-D
                1.5f, // Autorizado
                2.0f // Estado
        };
        final float[] W_APR = {
                0.8f, // Paso
                2.2f, // Departamento
                3.0f, // Autoriza
                1.6f, // Estado Flujo
                2.0f, // Fecha Autorización
                2.2f // Observación
        };

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
                    safe(request.getColorPrincipal())));

            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) {
                document.add(logo);
            }

            document.add(ReporteUtil.crearTituloEmpresa(safe(request.getEmpresa())));
            document.add(ReporteUtil.crearTituloReporte("REPORTE - SOLICITUDES DE VACACIONES"));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta())));

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
                if (emp == null)
                    continue;

                PdfPTable infoEmp = new PdfPTable(3);
                infoEmp.setWidthPercentage(100);
                infoEmp.setWidths(W_EMP_INFO);

                infoEmp.addCell(celdaInfoMixtaLocal("CIUDAD:", safe(emp.getCiudad()), Color.WHITE));
                infoEmp.addCell(celdaInfoMixtaLocal("C.C.:", safe(emp.getIdentificacion()), Color.WHITE));
                infoEmp.addCell(celdaInfoMixtaLocal("COD:", safe(emp.getCodigo()), Color.WHITE));

                PdfPCell cNombre = celdaInfoMixtaLocal("EMPLEADO:", safe(emp.getEmpleado()), Color.WHITE);
                cNombre.setColspan(3);
                infoEmp.addCell(cNombre);

                PdfPTable infoEmp2 = new PdfPTable(3);
                infoEmp2.setWidthPercentage(100);
                infoEmp2.setWidths(W_EMP_INFO_2);

                infoEmp2.addCell(celdaInfoMixtaLocal("RÉGIMEN LABORAL:", safe(emp.getRegimen()), Color.WHITE));
                infoEmp2.addCell(celdaInfoMixtaLocal("DEPARTAMENTO:", safe(emp.getDepartamento()), Color.WHITE));
                infoEmp2.addCell(celdaInfoMixtaLocal("CARGO:", safe(emp.getCargo()), Color.WHITE));

                PdfPTable contInfo = new PdfPTable(1);
                contInfo.setWidthPercentage(100);
                noSpace(contInfo);

                PdfPCell wrap1 = new PdfPCell(infoEmp);
                wrap1.setPadding(0f);
                wrap1.setBorder(Rectangle.BOX);
                wrap1.setBorderColor(new Color(120, 120, 120));
                contInfo.addCell(wrap1);

                PdfPCell wrap2 = new PdfPCell(infoEmp2);
                wrap2.setPadding(0f);
                wrap2.setBorder(Rectangle.BOX);
                wrap2.setBorderColor(new Color(120, 120, 120));
                contInfo.addCell(wrap2);

                contInfo.setSpacingBefore(6f);
                contInfo.setSpacingAfter(6f);
                document.add(contInfo);

PdfPTable tabla = new PdfPTable(7);
tabla.setWidthPercentage(100);
tabla.setWidths(W_SOL);

tabla.addCell(hCell("Solicitud", colorPrincipal));
tabla.addCell(hCell("Desde", colorPrincipal));
tabla.addCell(hCell("Hasta", colorPrincipal));
tabla.addCell(hCell("Días L-V", colorPrincipal));
tabla.addCell(hCell("Días S-D", colorPrincipal));
tabla.addCell(hCell("Autorizado", colorPrincipal));
tabla.addCell(hCell("Estado", colorPrincipal));

List<SolicitudVacacionReporteDTO> solicitudes =
        ordenarSolicitudes(emp.getSolicitudes());

if (solicitudes.isEmpty()) {

    PdfPCell sin = new PdfPCell(
            new Phrase(
                    "Sin solicitudes en el rango seleccionado",
                    ReporteUtil.fuenteTablaData()
            )
    );

    sin.setColspan(7);
    sin.setHorizontalAlignment(Element.ALIGN_CENTER);
    sin.setPadding(6f);

    tabla.addCell(sin);

} else {

    int i = 1;

    for (SolicitudVacacionReporteDTO s : solicitudes) {

        Color fondo = (i % 2 == 0)
                ? zebra
                : Color.WHITE;

        // =====================================================
        // DATOS DE LA SOLICITUD
        // =====================================================

        tabla.addCell(
                cellCenter(
                        String.valueOf(long0(s.getSolicitud())),
                        fondo
                )
        );

        tabla.addCell(
                cellCenter(
                        formatearFecha(s.getDesde()),
                        fondo
                )
        );

        tabla.addCell(
                cellCenter(
                        formatearFecha(s.getHasta()),
                        fondo
                )
        );

        tabla.addCell(
                cellCenter(
                        String.valueOf(num0(s.getDias_l_v())),
                        fondo
                )
        );

        tabla.addCell(
                cellCenter(
                        String.valueOf(num0(s.getDias_s_d())),
                        fondo
                )
        );

        tabla.addCell(
                cellCenter(
                        safe(s.getAutorizado()),
                        fondo
                )
        );

        tabla.addCell(
                cellCenter(
                        safe(s.getEstado_texto()),
                        fondo
                )
        );

        // =====================================================
        // HISTORIAL DE APROBACIONES
        // Solo se muestra cuando existe historial
        // =====================================================

        List<SolicitudVacacionAprobacionDTO> aprobaciones =
                s.getAprobaciones();

        if (aprobaciones != null && !aprobaciones.isEmpty()) {

            PdfPTable tablaAprobaciones = new PdfPTable(6);
            tablaAprobaciones.setWidthPercentage(100);
            tablaAprobaciones.setWidths(W_APR);

            tablaAprobaciones.addCell(
                    hCell("Paso", colorPrincipal)
            );

            tablaAprobaciones.addCell(
                    hCell("Departamento", colorPrincipal)
            );

            tablaAprobaciones.addCell(
                    hCell("Autoriza", colorPrincipal)
            );

            tablaAprobaciones.addCell(
                    hCell("Estado Flujo", colorPrincipal)
            );

            tablaAprobaciones.addCell(
                    hCell("Fecha Autorización", colorPrincipal)
            );

            tablaAprobaciones.addCell(
                    hCell("Observación", colorPrincipal)
            );

            for (SolicitudVacacionAprobacionDTO a : aprobaciones) {

                if (a == null) {
                    continue;
                }

                tablaAprobaciones.addCell(
                        cellCenter(
                                a.getOrden_paso() == null
                                        ? ""
                                        : String.valueOf(a.getOrden_paso()),
                                Color.WHITE
                        )
                );

                tablaAprobaciones.addCell(
                        cellLeft(
                                safe(a.getDepartamento_aprobacion()),
                                Color.WHITE
                        )
                );

                tablaAprobaciones.addCell(
                        cellLeft(
                                safe(a.getAutoriza()),
                                Color.WHITE
                        )
                );

                tablaAprobaciones.addCell(
                        cellCenter(
                                safe(a.getEstado_flujo()),
                                Color.WHITE
                        )
                );

                tablaAprobaciones.addCell(
                        cellCenter(
                                formatearFechaHora(
                                        a.getFecha_autorizacion()
                                ),
                                Color.WHITE
                        )
                );

                tablaAprobaciones.addCell(
                        cellLeft(
                                safe(a.getObservacion()).isBlank()
                                        ? "—"
                                        : safe(a.getObservacion()),
                                Color.WHITE
                        )
                );
            }

            PdfPCell historialCell =
                    new PdfPCell(tablaAprobaciones);

            historialCell.setColspan(7);
            historialCell.setPadding(0f);
            historialCell.setBorder(Rectangle.NO_BORDER);

            tabla.addCell(historialCell);
        }

        i++;
    }
}
                document.add(tabla);
            }

            document.close();
            return baos.toByteArray();

        }catch(

    IllegalArgumentException e)
    {
        throw e;
    }catch(
    Exception e)
    {
        throw new ReportBuildException("No se pudo generar ReporteSolicitudesVacaciones.pdf", e);
    }finally
    {
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
            CellStyle stCabGray = ConfiguracionExcel.crearEstiloCabeceraGris(libro);
            CellStyle stCabActivo = crearEstiloCabeceraEmpleadoActivo(libro, request.getColorPrincipal());

            XSSFSheet hoja = libro.createSheet("Solicitudes_Vacaciones");

            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) {
                insertarLogo(libro, hoja, logo);
            }

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

            mergeSafeNoBorder(hoja, row, row, 0, 8, estiloTitulo);
            UtilExcel.establecerTexto(hoja, row++, 0, "REPORTE - SOLICITUDES DE VACACIONES", estiloTitulo);

            mergeSafeNoBorder(hoja, row, row, 0, 8, estiloTitulo);
            UtilExcel.establecerTexto(hoja, row++, 0,
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta()),
                    estiloTitulo);

            row++;

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

                CellStyle stCab = stCabActivo;

                mergeSafeNoBorder(hoja, row, row, 0, 2, stCab);
                mergeSafeNoBorder(hoja, row, row, 3, 5, stCab);
                mergeSafeNoBorder(hoja, row, row, 6, 8, stCab);

                UtilExcel.establecerTexto(hoja, row, 0, "CIUDAD: " + safe(emp.getCiudad()), stCab);
                UtilExcel.establecerTexto(hoja, row, 3, "C.C.: " + safe(emp.getIdentificacion()), stCab);
                UtilExcel.establecerTexto(hoja, row, 6, "COD: " + safe(emp.getCodigo()), stCab);
                row++;

                mergeSafeNoBorder(hoja, row, row, 0, 8, stCab);
                UtilExcel.establecerTexto(hoja, row++, 0, "EMPLEADO: " + safe(emp.getEmpleado()), stCab);

                mergeSafeNoBorder(hoja, row, row, 0, 2, stCabGray);
                mergeSafeNoBorder(hoja, row, row, 3, 5, stCabGray);
                mergeSafeNoBorder(hoja, row, row, 6, 8, stCabGray);

                UtilExcel.establecerTexto(hoja, row, 0, "RÉGIMEN LABORAL: " + safe(emp.getRegimen()), stCabGray);
                UtilExcel.establecerTexto(hoja, row, 3, "DEPARTAMENTO: " + safe(emp.getDepartamento()), stCabGray);
                UtilExcel.establecerTexto(hoja, row, 6, "CARGO: " + safe(emp.getCargo()), stCabGray);
                row++;

                Row head = UtilExcel.asegurarFila(hoja, row++);
                UtilExcel.establecerTexto(head, 0, "Solicitud", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 1, "Desde", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 2, "Hasta", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 3, "Días L-V", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 4, "Días S-D", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 5, "Autorizado", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 6, "Estado", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 7, "Autoriza", estiloCentroBorde);
                UtilExcel.establecerTexto(head, 8, "Fecha de Autorización", estiloCentroBorde);

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

            hoja.createFreezePane(0, 4);

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
        if (filas == null || filas.isEmpty()) {
            return new ArrayList<>();
        }

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
                emp.setEmpleado(safe(row.getEmpleado()).isBlank()
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
                .sorted(Comparator
                        .comparing((SolicitudVacacionReporteDTO s) -> safe(s.getDesde()))
                        .thenComparing(s -> long0(s.getSolicitud())))
                .collect(Collectors.toList());
    }

    // =========================================================================================
    // Helpers PDF
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
        celda.setPadding(4f);
        celda.setBorder(Rectangle.NO_BORDER);
        return celda;
    }

    // =========================================================================================
    // Helpers Excel
    // =========================================================================================
    private void insertarLogo(Workbook wb, Sheet hoja, byte[] imagenBytes) {
        if (imagenBytes == null || imagenBytes.length == 0)
            return;

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

        short col = IndexedColors.GREY_50_PERCENT.getIndex();
        RegionUtil.setTopBorderColor(col, region, sh);
        RegionUtil.setBottomBorderColor(col, region, sh);
        RegionUtil.setLeftBorderColor(col, region, sh);
        RegionUtil.setRightBorderColor(col, region, sh);
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

    private void noSpace(PdfPTable t) {
        t.setSpacingBefore(0f);
        t.setSpacingAfter(0f);
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
