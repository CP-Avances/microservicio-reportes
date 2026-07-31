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
            if (request == null) {
                throw new IllegalArgumentException("request nulo");
            }

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
            document.add(ReporteUtil.crearTituloReporte("REPORTE DE SOLICITUDES DE PERMISOS"));
            document.add(ReporteUtil.crearTituloPeriodo(
                    "RANGO DE FECHAS: " + safe(request.getFechaDesde()) + " - " + safe(request.getFechaHasta())));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());

            List<SolicitudPermisoReporteDTO> filas = request.getSolicitudes();

            if (filas == null || filas.isEmpty()) {
                Paragraph p = new Paragraph("Sin datos para mostrar", ReporteUtil.fuenteTexto());
                p.setAlignment(Element.ALIGN_CENTER);
                p.setSpacingBefore(20f);
                document.add(p);
                document.close();
                return baos.toByteArray();
            }

            List<GrupoEmpleadoReporte> grupos = agruparPorEmpleadoSolicitud(filas);

            final float[] W_EMP_INFO = { 4f, 4f, 2f };

            for (GrupoEmpleadoReporte emp : grupos) {

                // ==========================
                // DATOS DEL EMPLEADO
                // ==========================
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

                // ==========================
                // SOLICITUDES DEL EMPLEADO
                // ==========================
                if (emp.solicitudes == null || emp.solicitudes.isEmpty()) {
                    Paragraph p = new Paragraph("Sin solicitudes registradas hasta la fecha de corte",
                            ReporteUtil.fuenteTexto());
                    p.setAlignment(Element.ALIGN_CENTER);
                    p.setSpacingBefore(8f);
                    p.setSpacingAfter(12f);
                    document.add(p);
                    continue;
                }

                for (GrupoSolicitudReporte sol : emp.solicitudes) {

                    // BLOQUE VISUAL POR CADA CÓDIGO DE SOLICITUD
                    PdfPTable bloqueSolicitud = new PdfPTable(1);
                    bloqueSolicitud.setWidthPercentage(100);
                    bloqueSolicitud.setSpacingBefore(6f);
                    bloqueSolicitud.setSpacingAfter(12f);

                    PdfPCell celdaBloque = new PdfPCell();
                    celdaBloque.setBorder(Rectangle.BOX);
                    celdaBloque.setBorderColor(new Color(70, 70, 70));
                    celdaBloque.setBorderWidth(0.8f);
                    celdaBloque.setPadding(0f);

                    celdaBloque.addElement(crearTablaCabeceraSolicitudPDF(sol, colorPrincipal));
                    celdaBloque.addElement(crearTablaHistorialSolicitudPDF(sol, colorPrincipal));

                    bloqueSolicitud.addCell(celdaBloque);
                    document.add(bloqueSolicitud);
                }
            }

            document.close();
            return baos.toByteArray();

        } catch (IllegalArgumentException e) {
            throw e;
        } catch (Exception e) {
            throw new ReportBuildException("No se pudo generar ReporteSolicitudesPermisos.pdf", e);
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
    // =========================================================================================
// EXCEL
// =========================================================================================
public byte[] generarReporteSolicitudesPermisosExcel(ReporteSolicitudesPermisosRequest request) {
    try (XSSFWorkbook libro = new XSSFWorkbook();
         ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

        if (request == null) {
            throw new IllegalArgumentException("request nulo");
        }

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

        // Columnas usadas para el nuevo formato
        sh.setColumnWidth(0, 4500);   // Código / Paso
        sh.setColumnWidth(1, 9000);   // Tipo permiso / Departamento
        sh.setColumnWidth(2, 8500);   // Fecha solicitud / Autoriza
        sh.setColumnWidth(3, 5000);   // Fecha inicio / Estado flujo
        sh.setColumnWidth(4, 5500);   // Fecha fin / Fecha autorización
        sh.setColumnWidth(5, 4500);   // Días / Observación
        sh.setColumnWidth(6, 3500);   // Horas
        sh.setColumnWidth(7, 4500);   // Autorizado
        sh.setColumnWidth(8, 5500);   // Estado solicitud

        int row = 0;
        int C0 = 0;
        int C_LAST = 8;

        // ==========================
        // TÍTULOS DEL REPORTE
        // ==========================
        mergeSafeNoBorder(sh, row, row, C0, C_LAST, estiloTitulo);
        UtilExcel.establecerTexto(
                sh,
                row++,
                C0,
                UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())),
                estiloTitulo
        );

        mergeSafeNoBorder(sh, row, row, C0, C_LAST, estiloTitulo);
        UtilExcel.establecerTexto(
                sh,
                row++,
                C0,
                "REPORTE DE SOLICITUDES DE PERMISOS",
                estiloTitulo
        );

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

        if (filas.isEmpty()) {
            mergeSafe(sh, row, row, C0, C_LAST, estiloCentroBorde);
            UtilExcel.establecerTexto(sh, row, C0, "Sin datos para mostrar", estiloCentroBorde);

            libro.write(baos);
            return baos.toByteArray();
        }

        List<GrupoEmpleadoReporte> grupos = agruparPorEmpleadoSolicitud(filas);

        for (GrupoEmpleadoReporte emp : grupos) {

            // ==========================
            // CABECERA DEL EMPLEADO
            // ==========================
            int rCabIni = row;

            mergeSafeNoBorder(sh, row, row, 0, 2, stCabActivo);
            mergeSafeNoBorder(sh, row, row, 3, 5, stCabActivo);
            mergeSafeNoBorder(sh, row, row, 6, 8, stCabActivo);

            UtilExcel.establecerTexto(sh, row, 0, "CIUDAD: " + safe(emp.ciudad), stCabActivo);
            UtilExcel.establecerTexto(sh, row, 3, "C.C.: " + safe(emp.identificacion), stCabActivo);
            UtilExcel.establecerTexto(sh, row, 6, "COD: " + safe(emp.codigo), stCabActivo);
            row++;

            mergeSafeNoBorder(sh, row, row, 0, 8, stCabActivo);
            UtilExcel.establecerTexto(sh, row, 0, "EMPLEADO: " + safe(emp.empleado), stCabActivo);
            row++;

            mergeSafeNoBorder(sh, row, row, 0, 2, stCabGray);
            mergeSafeNoBorder(sh, row, row, 3, 5, stCabGray);
            mergeSafeNoBorder(sh, row, row, 6, 8, stCabGray);

            UtilExcel.establecerTexto(sh, row, 0, "RÉGIMEN LABORAL: " + safe(emp.regimen), stCabGray);
            UtilExcel.establecerTexto(sh, row, 3, "DEPARTAMENTO: " + safe(emp.departamento), stCabGray);
            UtilExcel.establecerTexto(sh, row, 6, "CARGO: " + safe(emp.cargo), stCabGray);
            row++;

            int rCabFin = row - 1;
            limpiarBordesEnRegion(sh, rCabIni, rCabFin, 0, 8);
            bordeExternoGrueso(sh, new CellRangeAddress(rCabIni, rCabFin, 0, 8));

            row++;

            // ==========================
            // SOLICITUDES DEL EMPLEADO
            // ==========================
            if (emp.solicitudes == null || emp.solicitudes.isEmpty()) {
                mergeSafe(sh, row, row, 0, 8, estiloCentroBorde);
                UtilExcel.establecerTexto(
                        sh,
                        row++,
                        0,
                        "Sin solicitudes registradas hasta la fecha de corte",
                        estiloCentroBorde
                );

                row += 2;
                continue;
            }

            for (GrupoSolicitudReporte sol : emp.solicitudes) {
                int rBloqueIni = row;

                // ==========================
                // CABECERA DE LA SOLICITUD
                // ==========================
                Row hSol = UtilExcel.asegurarFila(sh, row++);
                UtilExcel.establecerTexto(hSol, 0, "Código Solicitud", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 1, "Tipo Permiso", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 2, "Fecha Solicitud", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 3, "Fecha Inicio", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 4, "Fecha Fin", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 5, "Días", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 6, "Horas", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 7, "Autorizado", stHeadBlue);
                UtilExcel.establecerTexto(hSol, 8, "Estado Solicitud", stHeadBlue);

                Row rSol = UtilExcel.asegurarFila(sh, row++);
                UtilExcel.establecerValor(rSol, 0, long0(sol.codigoSolicitud), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 1, safe(sol.tipoPermiso), estiloIzqBorde);
                UtilExcel.establecerTexto(rSol, 2, formatFechaHora(sol.fechaSolicitud), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 3, formatFecha(sol.fechaInicio), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 4, formatFecha(sol.fechaFin), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 5, fmtDec2(sol.dias), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 6, fmtDec2(sol.horas), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 7, safe(sol.autorizado), estiloCentroBorde);
                UtilExcel.establecerTexto(rSol, 8, safe(sol.estadoTexto), estiloCentroBorde);

                // ==========================
                // CABECERA DEL HISTORIAL
                // ==========================
                Row hHist = UtilExcel.asegurarFila(sh, row++);
                UtilExcel.establecerTexto(hHist, 0, "Paso", stCabGray);
                UtilExcel.establecerTexto(hHist, 1, "Departamento", stCabGray);
                UtilExcel.establecerTexto(hHist, 2, "Autoriza", stCabGray);
                UtilExcel.establecerTexto(hHist, 3, "Estado Flujo", stCabGray);
                UtilExcel.establecerTexto(hHist, 4, "Fecha Autorización", stCabGray);

                mergeSafe(sh, hHist.getRowNum(), hHist.getRowNum(), 5, 8, stCabGray);
                UtilExcel.establecerTexto(sh, hHist.getRowNum(), 5, "Observación", stCabGray);

                // ==========================
                // DETALLE DEL HISTORIAL
                // ==========================
                if (sol.aprobaciones == null || sol.aprobaciones.isEmpty()) {
                    mergeSafe(sh, row, row, 0, 8, estiloCentroBorde);
                    UtilExcel.establecerTexto(
                            sh,
                            row++,
                            0,
                            "Sin historial de aprobaciones registrado.",
                            estiloCentroBorde
                    );
                } else {
                    for (AprobacionSolicitudReporte a : sol.aprobaciones) {
                        Row rr = UtilExcel.asegurarFila(sh, row);

                        UtilExcel.establecerTexto(
                                rr,
                                0,
                                a.ordenPaso == null ? "" : String.valueOf(a.ordenPaso),
                                estiloCentroBorde
                        );

                        UtilExcel.establecerTexto(rr, 1, safe(a.departamentoAprobacion), estiloIzqBorde);
                        UtilExcel.establecerTexto(rr, 2, safe(a.autoriza), estiloIzqBorde);
                        UtilExcel.establecerTexto(rr, 3, safe(a.estadoFlujo), estiloCentroBorde);
                        UtilExcel.establecerTexto(rr, 4, formatFechaHora(a.fechaAutorizacion), estiloCentroBorde);

                        mergeSafe(sh, row, row, 5, 8, estiloIzqBorde);
                        UtilExcel.establecerTexto(
                                sh,
                                row,
                                5,
                                safe(a.observacion).isBlank() ? "—" : safe(a.observacion),
                                estiloIzqBorde
                        );

                        row++;
                    }
                }

                int rBloqueFin = row - 1;
                bordeExternoGrueso(sh, new CellRangeAddress(rBloqueIni, rBloqueFin, 0, 8));

                row++;
            }

            row++;
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
    // EMPLEADO -> SOLICITUD -> HISTORIAL DE APROBACIONES
    // =========================================================================================
    private List<GrupoEmpleadoReporte> agruparPorEmpleadoSolicitud(List<SolicitudPermisoReporteDTO> filas) {
        Map<Long, GrupoEmpleadoReporte> empleadosMap = new LinkedHashMap<>();

        List<SolicitudPermisoReporteDTO> ordenadas = (filas == null ? new ArrayList<SolicitudPermisoReporteDTO>()
                : filas)
                .stream()
                .filter(Objects::nonNull)
                .sorted(Comparator
                        .comparing((SolicitudPermisoReporteDTO x) -> safe(x.getApellido()))
                        .thenComparing(x -> safe(x.getNombre()))
                        .thenComparing(x -> long0(x.getCodigo_solicitud()))
                        .thenComparing(x -> x.getOrden_paso() == null ? 999 : x.getOrden_paso())
                        .thenComparing(x -> safe(x.getFecha_autorizacion()))
                        .thenComparing(x -> x.getId_historial() == null ? 0L : x.getId_historial()))
                .collect(Collectors.toList());

        for (SolicitudPermisoReporteDTO row : ordenadas) {
            Long idEmpleado = row.getId_empleado() != null ? row.getId_empleado() : 0L;

            GrupoEmpleadoReporte emp = empleadosMap.get(idEmpleado);

            if (emp == null) {
                emp = new GrupoEmpleadoReporte();
                emp.idEmpleado = idEmpleado;
                emp.identificacion = safe(row.getIdentificacion());
                emp.codigo = safe(row.getCodigo());
                emp.empleado = !safe(row.getEmpleado()).isBlank()
                        ? safe(row.getEmpleado())
                        : (safe(row.getApellido()) + " " + safe(row.getNombre())).trim();
                emp.ciudad = safe(row.getCiudad());
                emp.regimen = safe(row.getRegimen());
                emp.departamento = safe(row.getDepartamento());
                emp.cargo = safe(row.getCargo());
                emp.solicitudes = new ArrayList<>();

                empleadosMap.put(idEmpleado, emp);
            }

            Long idSolicitud = row.getSolicitud() != null ? row.getSolicitud() : 0L;

            GrupoSolicitudReporte sol = buscarSolicitud(emp.solicitudes, idSolicitud);

            if (sol == null) {
                sol = new GrupoSolicitudReporte();
                sol.solicitud = idSolicitud;
                sol.codigoSolicitud = row.getCodigo_solicitud();
                sol.idTipoPermiso = row.getId_tipo_permiso();
                sol.tipoPermiso = safe(row.getTipo_permiso());
                sol.fechaSolicitud = safe(row.getFecha_solicitud());
                sol.fechaInicio = safe(row.getFecha_inicio());
                sol.fechaFin = safe(row.getFecha_fin());
                sol.dias = row.getDias();
                sol.horas = row.getHoras();
                sol.minutosTotales = row.getMinutos_totales();
                sol.estado = row.getEstado();
                sol.estadoTexto = safe(row.getEstado_texto());
                sol.autorizado = safe(row.getAutorizado());
                sol.aprobaciones = new ArrayList<>();

                emp.solicitudes.add(sol);
            }

            if (row.getId_historial() != null) {
                boolean existeAprobacion = sol.aprobaciones.stream()
                        .anyMatch(a -> Objects.equals(a.idHistorial, row.getId_historial()));

                if (!existeAprobacion) {
                    AprobacionSolicitudReporte ap = new AprobacionSolicitudReporte();
                    ap.idHistorial = row.getId_historial();
                    ap.ordenPaso = row.getOrden_paso();
                    ap.idDepartamentoDestino = row.getId_departamento_destino();
                    ap.departamentoAprobacion = !safe(row.getDepartamento_aprobacion()).isBlank()
                            ? safe(row.getDepartamento_aprobacion())
                            : safe(row.getDepartamento_nombre());
                    ap.idEmpleadoAprobador = row.getId_empleado_aprobador();
                    ap.autoriza = !safe(row.getAutoriza()).isBlank()
                            ? safe(row.getAutoriza())
                            : safe(row.getEmpleado_nombre());
                    ap.accion = safe(row.getAccion());
                    ap.estadoFlujo = safe(row.getEstado_flujo());
                    ap.fechaAutorizacion = !safe(row.getFecha_autorizacion()).isBlank()
                            ? safe(row.getFecha_autorizacion())
                            : safe(row.getFecha_hora_accion());
                    ap.observacion = safe(row.getObservacion());
                    ap.tipoPaso = safe(row.getTipo_paso());
                    ap.obligatorio = row.getObligatorio();
                    ap.modoAprobador = safe(row.getModo_aprobador());
                    ap.cargoEnMomento = safe(row.getCargo_en_momento());
                    ap.esJefeEnMomento = row.getEs_jefe_en_momento();

                    sol.aprobaciones.add(ap);
                }
            }
        }

        for (GrupoEmpleadoReporte emp : empleadosMap.values()) {
            emp.solicitudes.sort(Comparator
                    .comparing((GrupoSolicitudReporte s) -> s.codigoSolicitud == null ? 0L : s.codigoSolicitud)
                    .thenComparing(s -> safe(s.fechaInicio))
                    .thenComparing(s -> s.solicitud == null ? 0L : s.solicitud));

            for (GrupoSolicitudReporte sol : emp.solicitudes) {
                sol.aprobaciones.sort(Comparator
                        .comparing((AprobacionSolicitudReporte a) -> a.ordenPaso == null ? 999 : a.ordenPaso)
                        .thenComparing(a -> safe(a.fechaAutorizacion))
                        .thenComparing(a -> a.idHistorial == null ? 0L : a.idHistorial));
            }
        }

        return new ArrayList<>(empleadosMap.values());
    }

    private GrupoSolicitudReporte buscarSolicitud(List<GrupoSolicitudReporte> solicitudes, Long idSolicitud) {
        if (solicitudes == null)
            return null;

        for (GrupoSolicitudReporte s : solicitudes) {
            if (Objects.equals(s.solicitud, idSolicitud)) {
                return s;
            }
        }

        return null;
    }

    private static class GrupoEmpleadoReporte {
        Long idEmpleado;
        String identificacion;
        String codigo;
        String empleado;
        String ciudad;
        String regimen;
        String departamento;
        String cargo;
        List<GrupoSolicitudReporte> solicitudes;
    }

    private static class GrupoSolicitudReporte {
        Long solicitud;
        Long codigoSolicitud;
        Long idTipoPermiso;
        String tipoPermiso;
        String fechaSolicitud;
        String fechaInicio;
        String fechaFin;
        Double dias;
        Double horas;
        Integer minutosTotales;
        Integer estado;
        String estadoTexto;
        String autorizado;
        List<AprobacionSolicitudReporte> aprobaciones;
    }

    private static class AprobacionSolicitudReporte {
        Long idHistorial;
        Integer ordenPaso;
        Long idDepartamentoDestino;
        String departamentoAprobacion;
        Long idEmpleadoAprobador;
        String autoriza;
        String accion;
        String estadoFlujo;
        String fechaAutorizacion;
        String observacion;
        String tipoPaso;
        Boolean obligatorio;
        String modoAprobador;
        String cargoEnMomento;
        Boolean esJefeEnMomento;
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
    PdfPCell c = new PdfPCell(new Phrase(safe(text), ReporteUtil.fuenteEncabezadoTablaData()));
    c.setHorizontalAlignment(Element.ALIGN_CENTER);
    c.setVerticalAlignment(Element.ALIGN_MIDDLE);
    c.setBackgroundColor(bg);
    c.setPadding(4f);
    c.setBorder(Rectangle.BOX);
    c.setBorderColor(new Color(80, 80, 80));
    c.setBorderWidth(0.8f);
    return c;
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

    private PdfPTable crearTablaCabeceraSolicitudPDF(GrupoSolicitudReporte sol, Color colorPrincipal)
            throws DocumentException {
        final float[] W_SOLICITUD = {
                1.5f, 3.2f, 2.2f, 1.8f, 1.8f, 0.9f, 0.9f, 1.3f, 1.8f
        };

        PdfPTable tabla = new PdfPTable(9);
        tabla.setWidthPercentage(100);
        tabla.setWidths(W_SOLICITUD);
        tabla.setSpacingAfter(0f);

        tabla.addCell(hCell("Código Solicitud", colorPrincipal));
        tabla.addCell(hCell("Tipo Permiso", colorPrincipal));
        tabla.addCell(hCell("Fecha Solicitud", colorPrincipal));
        tabla.addCell(hCell("Fecha Inicio", colorPrincipal));
        tabla.addCell(hCell("Fecha Fin", colorPrincipal));
        tabla.addCell(hCell("Días", colorPrincipal));
        tabla.addCell(hCell("Horas", colorPrincipal));
        tabla.addCell(hCell("Autorizado", colorPrincipal));
        tabla.addCell(hCell("Estado Solicitud", colorPrincipal));

        Color fondo = new Color(245, 245, 245);

        tabla.addCell(cellCenter(String.valueOf(long0(sol.codigoSolicitud)), fondo));
        tabla.addCell(cellLeft(safe(sol.tipoPermiso), fondo));
        tabla.addCell(cellCenter(formatFechaHora(sol.fechaSolicitud), fondo));
        tabla.addCell(cellCenter(formatFecha(sol.fechaInicio), fondo));
        tabla.addCell(cellCenter(formatFecha(sol.fechaFin), fondo));
        tabla.addCell(cellCenter(fmtDec2(sol.dias), fondo));
        tabla.addCell(cellCenter(fmtDec2(sol.horas), fondo));
        tabla.addCell(cellCenter(safe(sol.autorizado), fondo));
        tabla.addCell(cellCenter(safe(sol.estadoTexto), fondo));

        return tabla;
    }

    private PdfPTable crearTablaHistorialSolicitudPDF(GrupoSolicitudReporte sol, Color colorPrincipal)
            throws DocumentException {
        final float[] W_HISTORIAL = {
                0.8f, 2.6f, 3.2f, 1.8f, 2.3f, 4.3f
        };

        PdfPTable tabla = new PdfPTable(6);
        tabla.setWidthPercentage(100);
        tabla.setWidths(W_HISTORIAL);
        tabla.setSpacingBefore(0f);
        tabla.setSpacingAfter(0f);

        Color fondoHeader = new Color(250, 250, 250);

        tabla.addCell(hCell("Paso", fondoHeader));
        tabla.addCell(hCell("Departamento", fondoHeader));
        tabla.addCell(hCell("Autoriza", fondoHeader));
        tabla.addCell(hCell("Estado Flujo", fondoHeader));
        tabla.addCell(hCell("Fecha Autorización", fondoHeader));
        tabla.addCell(hCell("Observación", fondoHeader));

        if (sol.aprobaciones == null || sol.aprobaciones.isEmpty()) {
            PdfPCell sinHistorial = cellCenterColspan("Sin historial de aprobaciones registrado.", Color.WHITE, 6);
            tabla.addCell(sinHistorial);
            return tabla;
        }

        int i = 1;
        Color zebra = ReporteUtil.colorZebraClaro();

        for (AprobacionSolicitudReporte a : sol.aprobaciones) {
            Color fondo = (i % 2 == 0) ? zebra : Color.WHITE;

            tabla.addCell(cellCenter(a.ordenPaso == null ? "" : String.valueOf(a.ordenPaso), fondo));
            tabla.addCell(cellLeft(safe(a.departamentoAprobacion), fondo));
            tabla.addCell(cellLeft(safe(a.autoriza), fondo));
            tabla.addCell(cellCenter(safe(a.estadoFlujo), fondo));
            tabla.addCell(cellCenter(formatFechaHora(a.fechaAutorizacion), fondo));
            tabla.addCell(cellLeft(safe(a.observacion).isBlank() ? "—" : safe(a.observacion), fondo));

            i++;
        }

        return tabla;
    }

    private PdfPCell cellCenterColspan(String text, Color bg, int colspan) {
        PdfPCell c = new PdfPCell(new Phrase(safe(text), ReporteUtil.fuenteTablaData()));
        c.setHorizontalAlignment(Element.ALIGN_CENTER);
        c.setVerticalAlignment(Element.ALIGN_MIDDLE);
        c.setBackgroundColor(bg);
        c.setPadding(4f);
        c.setColspan(colspan);
        c.setBorder(Rectangle.BOX);
        c.setBorderColor(new Color(80, 80, 80));
        c.setBorderWidth(0.6f);
        return c;
    }

    // =========================================================================================
    // HELPERS EXCEL
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

    private void limpiarBordesEnRegion(XSSFSheet sh, int r1, int r2, int c1, int c2) {
        int rr1 = Math.min(r1, r2), rr2 = Math.max(r1, r2);
        int cc1 = Math.min(c1, c2), cc2 = Math.max(c1, c2);

        for (int r = rr1; r <= rr2; r++) {
            Row row = UtilExcel.asegurarFila(sh, r);
            for (int c = cc1; c <= cc2; c++) {
                Cell cell = row.getCell(c);
                if (cell == null)
                    cell = row.createCell(c);

                CellStyle st = cell.getCellStyle();
                if (st == null)
                    continue;

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
    // HELPERS GENERALES
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

    private String fmtDec2(Double n) {
        if (n == null)
            return "0,00";
        return String.format(Locale.US, "%.2f", n).replace('.', ',');
    }

    private String formatFecha(String fecha) {
        String v = safe(fecha);
        if (v.isBlank())
            return "";
        if (v.length() >= 10 && v.matches("^\\d{4}-\\d{2}-\\d{2}.*")) {
            return v.substring(8, 10) + "/" + v.substring(5, 7) + "/" + v.substring(0, 4);
        }
        return v;
    }

    private String formatFechaHora(String fecha) {
        String v = safe(fecha);
        if (v.isBlank())
            return "";

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