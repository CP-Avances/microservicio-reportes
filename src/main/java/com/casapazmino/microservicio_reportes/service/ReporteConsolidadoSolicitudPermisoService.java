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
            if (request == null) throw new IllegalArgumentException("request nulo");
            baos = new ByteArrayOutputStream();
            document = new Document(PageSize.A4.rotate(), 30, 30, 30, 50);
            writer = PdfWriter.getInstance(document, baos);
            writer.setPageEvent(new ConfiguracionPaginaPDF(safe(request.getUsuario()), safe(request.getFraseMarcaAgua()), safe(request.getColorPrincipal())));
            document.open();

            Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
            if (logo != null) document.add(logo);

            String tipoUsuarios = obtenerTextoUsuarios(request.getUsuarios());
            document.add(ReporteUtil.crearTituloEmpresa(safe(request.getEmpresa())));
            document.add(ReporteUtil.crearTituloReporte("REPORTE DE SOLICITUDES DE PERMISOS - USUARIOS " + tipoUsuarios));
            document.add(ReporteUtil.crearTituloPeriodo("PERIODO DEL: " + safe(request.getFechaDesde()) + " AL " + safe(request.getFechaHasta())));

            Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
            Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
            Color zebraColor = ReporteUtil.colorZebraClaro();

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
            int totalSolicitudes = grupos.stream().mapToInt(g -> g.solicitudes == null ? 0 : g.solicitudes.size()).sum();

            PdfPTable tituloTabla = new PdfPTable(2);
            tituloTabla.setWidthPercentage(100);
            tituloTabla.setWidths(new float[]{8f, 2f});
            tituloTabla.setSpacingAfter(10f);

            PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA DE SOLICITUDES", ReporteUtil.fuenteEncabezadoTablaData()));
            celdaTitulo.setBackgroundColor(colorSecundario);
            celdaTitulo.setPadding(5f);
            celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
            tituloTabla.addCell(celdaTitulo);

            PdfPCell celdaContador = new PdfPCell(new Phrase("N° Registros: " + totalSolicitudes, ReporteUtil.fuenteEncabezadoTablaData()));
            celdaContador.setBackgroundColor(colorSecundario);
            celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
            celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
            celdaContador.setPadding(5f);
            celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
            tituloTabla.addCell(celdaContador);
            document.add(tituloTabla);

            int contadorSolicitud = 0;
            for (GrupoEmpleadoReporte emp : grupos) {
                PdfPTable infoEmpleado = new PdfPTable(3);
                infoEmpleado.setWidthPercentage(100);
                infoEmpleado.setWidths(new float[]{4f, 4f, 4f});

                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.identificacion, zebraColor));
                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:", emp.empleado, zebraColor));
                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.codigo, zebraColor));
                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CIUDAD:", emp.ciudad, zebraColor));
                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.regimen, zebraColor));
                infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.departamento, zebraColor));

                PdfPCell cargo = ReporteUtil.celdaInfoMixta("CARGO:", emp.cargo, zebraColor);
                cargo.setColspan(3);
                infoEmpleado.addCell(cargo);

                PdfPTable contenedora = new PdfPTable(1);
                contenedora.setWidthPercentage(100);
                contenedora.setSpacingAfter(7f);
                PdfPCell contenedor = new PdfPCell(infoEmpleado);
                contenedor.setPadding(0f);
                contenedor.setBorder(Rectangle.BOX);
                contenedora.addCell(contenedor);
                document.add(contenedora);

                if (emp.solicitudes == null || emp.solicitudes.isEmpty()) {
                    Paragraph p = new Paragraph("Sin solicitudes registradas hasta la fecha de corte", ReporteUtil.fuenteTexto());
                    p.setAlignment(Element.ALIGN_CENTER);
                    p.setSpacingAfter(10f);
                    document.add(p);
                    continue;
                }

                for (GrupoSolicitudReporte sol : emp.solicitudes) {
                    Color fondo = contadorSolicitud++ % 2 == 0 ? Color.WHITE : zebraColor;
                    PdfPTable bloque = new PdfPTable(1);
                    bloque.setWidthPercentage(100);
                    bloque.setSpacingAfter(10f);
                    bloque.setKeepTogether(true);

                    PdfPCell contenido = new PdfPCell();
                    contenido.setBorder(Rectangle.NO_BORDER);
                    contenido.setPadding(0f);
                    contenido.addElement(crearTablaCabeceraSolicitudPDF(sol, colorPrincipal, colorSecundario, fondo));
                    contenido.addElement(crearTablaHistorialSolicitudPDF(sol, colorSecundario));
                    bloque.addCell(contenido);
                    document.add(bloque);
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
        final String NOMBRE_HOJA = "Solicitudes_Permisos";
        final int FILA_ENCABEZADO = 5;
        final String[] HEADERS = {
                "ITEM","IDENTIFICACIÓN","CÓDIGO","APELLIDO NOMBRE","CIUDAD","SUCURSAL","RÉGIMEN","DEPARTAMENTO","CARGO",
                "CÓDIGO SOLICITUD","TIPO PERMISO","FECHA SOLICITUD","FECHA INICIO","FECHA FIN","HORA INICIO","TIMBRE INICIO",
                "HORA FIN","TIMBRE FIN","DÍAS","DURACIÓN","AUTORIZADO","ESTADO SOLICITUD","PASO","DEPARTAMENTO APROBACIÓN",
                "AUTORIZA","ESTADO FLUJO","FECHA AUTORIZACIÓN","OBSERVACIÓN"
        };
        final int[] ANCHOS = {10,20,15,28,18,24,24,28,22,18,28,22,18,18,16,18,16,18,10,14,14,20,10,28,28,18,22,30};

        try (XSSFWorkbook libro = new XSSFWorkbook(); ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            if (request == null) throw new IllegalArgumentException("request nulo");
            XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
            hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

            byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
            if (logo != null && logo.length > 0) UtilExcel.insertarLogoEstandar(libro, hoja, logo);

            for (int fila = 0; fila <= 4; fila++) UtilExcel.combinarCeldas(hoja, fila, fila, 1, HEADERS.length - 1);

            CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
            UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 1, 1, "REPORTE DE SOLICITUDES DE PERMISOS - USUARIOS " + obtenerTextoUsuarios(request.getUsuarios()), estiloTitulo);
            UtilExcel.establecerTexto(hoja, 2, 1, "PERIODO DEL REPORTE: " + safe(request.getFechaDesde()) + " AL " + safe(request.getFechaHasta()), estiloTitulo);

            Row encabezado = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
            for (int columna = 0; columna < HEADERS.length; columna++) UtilExcel.establecerTexto(encabezado, columna, HEADERS[columna], null);
            CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
            UtilExcel.aplicarEstiloAFila(encabezado, HEADERS.length, estiloHeader);
            UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
            encabezado.setHeightInPoints(18f);

            List<GrupoEmpleadoReporte> grupos = agruparPorEmpleadoSolicitud(request.getSolicitudes());
            int filaActual = FILA_ENCABEZADO + 1;
            int item = 1;

            for (GrupoEmpleadoReporte empleado : grupos) {
                for (GrupoSolicitudReporte solicitud : empleado.solicitudes) {
                    List<AprobacionSolicitudReporte> aprobaciones = solicitud.aprobaciones == null || solicitud.aprobaciones.isEmpty()
                            ? Collections.singletonList(null) : solicitud.aprobaciones;
                    boolean permisoPorHoras = esPermisoPorHoras(solicitud);

                    for (AprobacionSolicitudReporte aprobacion : aprobaciones) {
                        Row fila = UtilExcel.asegurarFila(hoja, filaActual++);
                        int columna = 0;
                        UtilExcel.establecerValor(fila, columna++, item++, null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.identificacion), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.codigo), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.empleado), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.ciudad), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.sucursal), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.regimen), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.departamento), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(empleado.cargo), null);
                        UtilExcel.establecerValor(fila, columna++, long0(solicitud.codigoSolicitud), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(solicitud.tipoPermiso), null);
                        UtilExcel.establecerTexto(fila, columna++, formatFechaHora(solicitud.fechaSolicitud), null);
                        UtilExcel.establecerTexto(fila, columna++, formatFecha(solicitud.fechaInicio), null);
                        UtilExcel.establecerTexto(fila, columna++, formatFecha(solicitud.fechaFin), null);
                        UtilExcel.establecerTexto(fila, columna++, permisoPorHoras ? formatearHora(solicitud.horaInicio) : "", null);
                        UtilExcel.establecerTexto(fila, columna++, permisoPorHoras ? formatearHora(solicitud.timbreInicioPermiso) : "", null);
                        UtilExcel.establecerTexto(fila, columna++, permisoPorHoras ? formatearHora(solicitud.horaFin) : "", null);
                        UtilExcel.establecerTexto(fila, columna++, permisoPorHoras ? formatearHora(solicitud.timbreFinPermiso) : "", null);
                        UtilExcel.establecerTexto(fila, columna++, formatearDias(solicitud.dias), null);
                        UtilExcel.establecerTexto(fila, columna++, permisoPorHoras ? formatearDuracion(solicitud.minutosTotales) : "", null);
                        UtilExcel.establecerTexto(fila, columna++, safe(solicitud.autorizado), null);
                        UtilExcel.establecerTexto(fila, columna++, safe(solicitud.estadoTexto), null);
                        UtilExcel.establecerTexto(fila, columna++, aprobacion == null || aprobacion.ordenPaso == null ? "" : String.valueOf(aprobacion.ordenPaso), null);
                        UtilExcel.establecerTexto(fila, columna++, aprobacion == null ? "" : safe(aprobacion.departamentoAprobacion), null);
                        UtilExcel.establecerTexto(fila, columna++, aprobacion == null ? "" : safe(aprobacion.autoriza), null);
                        UtilExcel.establecerTexto(fila, columna++, aprobacion == null ? "" : safe(aprobacion.estadoFlujo), null);
                        UtilExcel.establecerTexto(fila, columna++, aprobacion == null ? "" : formatFechaHora(aprobacion.fechaAutorizacion), null);
                        UtilExcel.establecerTexto(fila, columna, aprobacion == null ? "" : safe(aprobacion.observacion), null);
                    }
                }
            }

            int ultimaFila = filaActual == FILA_ENCABEZADO + 1 ? FILA_ENCABEZADO : filaActual - 1;
            CellStyle estiloCentro = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
            CellStyle estiloIzquierda = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);
            UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1, estiloCentro, true);

            if (ultimaFila > FILA_ENCABEZADO) {
                UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO + 1, ultimaFila, 0, HEADERS.length - 1, estiloCentro, true);
                int[] columnasIzquierda = {3,5,6,7,8,10,23,24,27};
                for (int columna : columnasIzquierda) UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO + 1, ultimaFila, columna, columna, estiloIzquierda, true);

                boolean[] filtros = new boolean[HEADERS.length];
                Arrays.fill(filtros, true);
                filtros[0] = false;
                UtilExcel.crearTablaEstilizada(hoja, "SolicitudesPermisosReporteTabla", FILA_ENCABEZADO, 0, ultimaFila, HEADERS.length - 1, true, filtros);
            }

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
                emp.ciudad = safe(row.getCiudad());
                emp.sucursal = safe(row.getSucursal());
                emp.regimen = safe(row.getRegimen());
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
                sol.horaInicio = safe(row.getHora_inicio());
                sol.timbreInicioPermiso = safe(row.getTimbre_inicio_permiso());
                sol.horaFin = safe(row.getHora_fin());
                sol.timbreFinPermiso = safe(row.getTimbre_fin_permiso());
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
        String sucursal;
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
        String horaInicio;
        String timbreInicioPermiso;
        String horaFin;
        String timbreFinPermiso;
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


    private PdfPTable crearTablaCabeceraSolicitudPDF(GrupoSolicitudReporte sol, Color colorPrincipal, Color colorSecundario, Color fondo) throws DocumentException {
        final float[] widths = {0.8f, 2.5f, 1.8f, 1.4f, 1.4f, 1.1f, 1.2f, 1.1f, 1.2f, 0.7f, 1.0f, 1.2f, 1.5f};
        PdfPTable tabla = new PdfPTable(13);
        tabla.setWidthPercentage(100);
        tabla.setWidths(widths);
        tabla.setSpacingAfter(0f);

        tabla.addCell(ReporteUtil.crearCelda("CÓDIGO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
        tabla.addCell(ReporteUtil.crearCelda("TIPO PERMISO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
        tabla.addCell(ReporteUtil.crearCelda("FECHA SOLICITUD", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
        tabla.addCell(ReporteUtil.crearCelda("PERIODO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));
        tabla.addCell(ReporteUtil.crearCelda("INICIO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));
        tabla.addCell(ReporteUtil.crearCelda("FIN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));
        tabla.addCell(ReporteUtil.crearCelda("TIEMPO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 1, 2));
        tabla.addCell(ReporteUtil.crearCelda("AUTORIZADO", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));
        tabla.addCell(ReporteUtil.crearCelda("ESTADO SOLICITUD", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal, 2, 1));

        tabla.addCell(ReporteUtil.crearCelda("DESDE", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
        tabla.addCell(ReporteUtil.crearCelda("HASTA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
        tabla.addCell(ReporteUtil.crearCelda("SOLICITADA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
        tabla.addCell(ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("SOLICITADA", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
        tabla.addCell(ReporteUtil.crearCelda("TIMBRE", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("DÍAS", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));
        tabla.addCell(ReporteUtil.crearCelda("DURACIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorPrincipal));

        boolean permisoPorHoras = esPermisoPorHoras(sol);
        tabla.addCell(cellCenter(String.valueOf(long0(sol.codigoSolicitud)), fondo));
        tabla.addCell(cellLeft(safe(sol.tipoPermiso), fondo));
        tabla.addCell(cellCenter(formatFechaHora(sol.fechaSolicitud), fondo));
        tabla.addCell(cellCenter(formatFecha(sol.fechaInicio), fondo));
        tabla.addCell(cellCenter(formatFecha(sol.fechaFin), fondo));
        tabla.addCell(cellCenter(permisoPorHoras ? formatearHora(sol.horaInicio) : "", fondo));
        tabla.addCell(cellCenter(permisoPorHoras ? formatearHora(sol.timbreInicioPermiso) : "", fondo));
        tabla.addCell(cellCenter(permisoPorHoras ? formatearHora(sol.horaFin) : "", fondo));
        tabla.addCell(cellCenter(permisoPorHoras ? formatearHora(sol.timbreFinPermiso) : "", fondo));
        tabla.addCell(cellCenter(formatearDias(sol.dias), fondo));
        tabla.addCell(cellCenter(permisoPorHoras ? formatearDuracion(sol.minutosTotales) : "", fondo));
        tabla.addCell(cellCenter(safe(sol.autorizado), fondo));
        tabla.addCell(cellCenter(safe(sol.estadoTexto), fondo));
        return tabla;
    }

    private PdfPTable crearTablaHistorialSolicitudPDF(GrupoSolicitudReporte sol, Color colorSecundario) throws DocumentException {
        final float[] widths = {0.8f, 2.6f, 3.2f, 1.8f, 2.3f, 4.3f};
        PdfPTable tabla = new PdfPTable(6);
        tabla.setWidthPercentage(100);
        tabla.setWidths(widths);
        tabla.setSpacingBefore(0f);
        tabla.setSpacingAfter(0f);

        tabla.addCell(ReporteUtil.crearCelda("PASO", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("DEPARTAMENTO", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("AUTORIZA", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("ESTADO FLUJO", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("FECHA AUTORIZACIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));
        tabla.addCell(ReporteUtil.crearCelda("OBSERVACIÓN", ReporteUtil.fuenteEncabezadoTablaData(), colorSecundario));

        if (sol.aprobaciones == null || sol.aprobaciones.isEmpty()) {
            tabla.addCell(cellCenterColspan("Sin historial de aprobaciones registrado.", Color.WHITE, 6));
            return tabla;
        }

        int contador = 1;
        Color zebra = ReporteUtil.colorZebraClaro();
        for (AprobacionSolicitudReporte aprobacion : sol.aprobaciones) {
            Color fondo = contador++ % 2 == 0 ? zebra : Color.WHITE;
            tabla.addCell(cellCenter(aprobacion.ordenPaso == null ? "" : String.valueOf(aprobacion.ordenPaso), fondo));
            tabla.addCell(cellLeft(safe(aprobacion.departamentoAprobacion), fondo));
            tabla.addCell(cellLeft(safe(aprobacion.autoriza), fondo));
            tabla.addCell(cellCenter(safe(aprobacion.estadoFlujo), fondo));
            tabla.addCell(cellCenter(formatFechaHora(aprobacion.fechaAutorizacion), fondo));
            tabla.addCell(cellLeft(safe(aprobacion.observacion).isBlank() ? "—" : safe(aprobacion.observacion), fondo));
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

    private String obtenerTextoUsuarios(Map<String, Boolean> usuarios) {
        if (usuarios == null) return "ACTIVOS";
        boolean activos = Boolean.TRUE.equals(usuarios.get("activos"));
        boolean inactivos = Boolean.TRUE.equals(usuarios.get("inactivos"));
        if (activos && inactivos) return "ACTIVOS E INACTIVOS";
        if (inactivos) return "INACTIVOS";
        return "ACTIVOS";
    }

    private String formatearDias(Double dias) {
        if (dias == null) return "0";
        if (dias.doubleValue() == Math.rint(dias.doubleValue())) return String.valueOf(dias.intValue());
        return String.format(Locale.US, "%.2f", dias).replace('.', ',');
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


    private boolean esPermisoPorHoras(GrupoSolicitudReporte sol) {
        if (sol == null) {
            return false;
        }

        double dias = sol.dias == null ? 0D : sol.dias;

        return Double.compare(dias, 0D) == 0
                && !safe(sol.horaInicio).isBlank()
                && !safe(sol.horaFin).isBlank();
    }

    private String formatearHora(String valor) {
        String hora = safe(valor);

        if (hora.isBlank()) {
            return "";
        }

        int posicionDosPuntos = hora.indexOf(':');

        if (posicionDosPuntos >= 2 && posicionDosPuntos + 2 < hora.length()) {
            return hora.substring(posicionDosPuntos - 2, posicionDosPuntos + 3);
        }

        return hora;
    }

    private String formatearDuracion(Integer minutosTotales) {
        int total = minutosTotales == null ? 0 : Math.max(0, minutosTotales);
        int horas = total / 60;
        int minutos = total % 60;

        return String.format("%02d:%02d", horas, minutos);
    }
    
}