package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ResumenAsistencia.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openpdf.text.Document;
import org.openpdf.text.Element;
import org.openpdf.text.Font;
import org.openpdf.text.Image;
import org.openpdf.text.PageSize;
import org.openpdf.text.Phrase;
import org.openpdf.text.Rectangle;
import org.openpdf.text.pdf.PdfPCell;
import org.openpdf.text.pdf.PdfPTable;
import org.openpdf.text.pdf.PdfWriter;
import org.springframework.stereotype.Service;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.concurrent.atomic.AtomicInteger;

@Service
public class ReporteAsistenciaService {

  public byte[] generarReporteResumenAsistenciaPDF(ReporteAsistenciaRequest request) {

    final float[] WIDTHS_7 = { 2.2f, 2.2f, 2.2f, 2.6f, 2.2f, 2.2f, 2.2f };
    final float[] WIDTHS_2 = { 8, 2 };

    final float[] WIDTHS_22 = {
        0.6f, 1.4f,
        1.30f, 1.30f, 0.65f,
        1.30f, 1.30f, 0.65f,
        1.30f, 1.30f, 0.65f,
        1.30f, 1.30f, 0.65f,
        1.1f, 1.7f,
        1.35f, 1.35f, 1.35f,
        1.7f, 1.5f, 2.8f
    };

    final Color COLOR_FALTA_TIMBRE = new Color(0xE05555); // rojo ligeramente suavizado
    final Color COLOR_ATRASO = new Color(0xE6DA55); // amarillo menos chillón
    final Color COLOR_SALIDA_ANTICIPADA = new Color(0x5AA3E6); // azul un poco más suave
    final Color COLOR_EXCESO_ALIMENTACION = new Color(0x66E055); // verde menos intenso
    final Color COLOR_PERMISO = new Color(0xD1A15A); // naranja/café apenas suavizado
    final Color COLOR_VACACIONES = new Color(0xA995D1); // lila un poco menos saturado
    final Color COLOR_JUSTIFICACION_HORAS_EXTRAS = new Color(0x4DB6AC);
    final Font fontHeaderCompacto = ReporteUtil.fuenteEncabezadoCompacto();
    final Font fontTextoCompacto = ReporteUtil.fuenteTextoCompacto();

    Document document = null;
    PdfWriter writer = null;
    ByteArrayOutputStream baos = null;

    try {
      baos = new ByteArrayOutputStream();
      document = new Document(PageSize.A4.rotate(), 30, 30, 30, 50);
      writer = PdfWriter.getInstance(document, baos);
      writer.setPageEvent(new ConfiguracionPaginaPDF(
          request.getUsuario(),
          request.getFraseMarcaAgua(),
          request.getColorPrincipal()));
      document.open();

      Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
      if (logo != null) {
        document.add(logo);
      }

      document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));
      String titulo = "RESUMEN DE ASISTENCIA - "
          + (request.getOpcionBusqueda() != null && request.getOpcionBusqueda() == 1 ? "ACTIVOS" : "INACTIVOS");
      document.add(ReporteUtil.crearTituloReporte(titulo));
      document.add(ReporteUtil.crearTituloPeriodo(
          "PERIODO DEL: " + safe(request.getFechaInicio()) + " AL " + safe(request.getFechaFin())));

      Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
      Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
      Color zebraColor = ReporteUtil.colorZebraClaro();

      PdfPTable colores = new PdfPTable(7);
      colores.setWidthPercentage(100);
      colores.setWidths(WIDTHS_7);
      colores.addCell(ReporteUtil.celdaEncabezado("CÓDIGO DE COLOR", Color.WHITE));
      colores.addCell(ReporteUtil.celdaEncabezado("FALTA TIMBRE", COLOR_FALTA_TIMBRE));
      colores.addCell(ReporteUtil.celdaEncabezado("ATRASO", COLOR_ATRASO));
      colores.addCell(ReporteUtil.celdaEncabezado("SALIDA ANTICIPADA", COLOR_SALIDA_ANTICIPADA));
      colores.addCell(ReporteUtil.celdaEncabezado("EXCESO DE ALIMENTACIÓN", COLOR_EXCESO_ALIMENTACION));
      colores.addCell(ReporteUtil.celdaEncabezado("PERMISO", COLOR_PERMISO));
      colores.addCell(ReporteUtil.celdaEncabezado("VACACIONES", COLOR_VACACIONES));
      colores.setSpacingAfter(10f);
      document.add(colores);

      AtomicInteger totalRegistros = new AtomicInteger();
      if (request.getGrupos() != null) {
        request.getGrupos().forEach(grupo -> {
          if (grupo.getEmpleados() != null) {
            grupo.getEmpleados().forEach(emp -> {
              totalRegistros.addAndGet(emp.getTLaborado() != null ? emp.getTLaborado().size() : 0);
            });
          }
        });
      }

      PdfPTable tituloTabla = new PdfPTable(2);
      tituloTabla.setWidthPercentage(100);
      tituloTabla.setWidths(WIDTHS_2);
      tituloTabla.setSpacingAfter(5f);

      PdfPCell celdaTitulo = new PdfPCell(new Phrase("LISTA EMPLEADOS", ReporteUtil.fuenteEncabezado()));
      celdaTitulo.setBackgroundColor(colorSecundario);
      celdaTitulo.setPadding(5f);
      celdaTitulo.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.LEFT);
      tituloTabla.addCell(celdaTitulo);

      PdfPCell celdaContador = new PdfPCell(
          new Phrase("N° Registros: " + totalRegistros.get(), ReporteUtil.fuenteEncabezado()));
      celdaContador.setBackgroundColor(colorSecundario);
      celdaContador.setHorizontalAlignment(Element.ALIGN_RIGHT);
      celdaContador.setVerticalAlignment(Element.ALIGN_MIDDLE);
      celdaContador.setPadding(5f);
      celdaContador.setBorder(Rectangle.TOP | Rectangle.BOTTOM | Rectangle.RIGHT);
      tituloTabla.addCell(celdaContador);

      tituloTabla.setSpacingAfter(10f);
      document.add(tituloTabla);

      if (request.getGrupos() != null) {
        for (GrupoAsistenciaDTO grupo : request.getGrupos()) {
          if (grupo.getEmpleados() == null) {
            continue;
          }

          for (EmpleadoAsistenciaDTO emp : grupo.getEmpleados()) {
            double totalAtrasos = 0;
            double totalSalidasAnticipadas = 0;
            double totalAlimentacionTomado = 0;
            double totalAlimentacionAsignado = 0;
            double totalExcesoAlimentacion = 0;
            double totalPlanificado = 0;
            double totalLaborado = 0;

            PdfPTable infoEmpleado = new PdfPTable(3);
            infoEmpleado.setWidthPercentage(100);
            infoEmpleado.setWidths(new float[] { 4, 4, 4 });

            infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("C.C.:", emp.getIdentificacion(), zebraColor));
            infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("EMPLEADO:",
                safe(emp.getApellido()) + " " + safe(emp.getNombre()), zebraColor));
            infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("COD:", emp.getCodigo(), zebraColor));
            infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN LABORAL:", emp.getRegimen(), zebraColor));
            infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("DEPARTAMENTO:", emp.getDepartamento(), zebraColor));
            infoEmpleado.addCell(ReporteUtil.celdaInfoMixta("CARGO:", emp.getCargo(), zebraColor));

            PdfPTable tablaContenedora = new PdfPTable(1);
            tablaContenedora.setWidthPercentage(100);
            PdfPCell contenedor = new PdfPCell(infoEmpleado);
            contenedor.setPadding(0);
            contenedor.setBorder(Rectangle.BOX);
            tablaContenedora.addCell(contenedor);
            tablaContenedora.setSpacingAfter(5f);
            document.add(tablaContenedora);

            PdfPTable encabezado = new PdfPTable(22);
            encabezado.setWidthPercentage(100);
            encabezado.setWidths(WIDTHS_22);

            encabezado.addCell(ReporteUtil.crearCeldaCompacta("N°", fontHeaderCompacto, colorPrincipal, 2, 1));
            encabezado.addCell(ReporteUtil.crearCeldaCompacta("FECHA", fontHeaderCompacto, colorPrincipal, 2, 1));

            encabezado.addCell(ReporteUtil.crearCeldaCompacta("ENTRADA", fontHeaderCompacto, colorPrincipal, 1, 3));
            encabezado.addCell(
                ReporteUtil.crearCeldaCompacta("INICIO ALIMENTACIÓN", fontHeaderCompacto, colorPrincipal, 1, 3));
            encabezado
                .addCell(ReporteUtil.crearCeldaCompacta("FIN ALIMENTACIÓN", fontHeaderCompacto, colorPrincipal, 1, 3));
            encabezado.addCell(ReporteUtil.crearCeldaCompacta("SALIDA", fontHeaderCompacto, colorPrincipal, 1, 3));

            encabezado.addCell(ReporteUtil.crearCeldaCompacta("ATRASO", fontHeaderCompacto, colorPrincipal, 2, 1));
            encabezado
                .addCell(ReporteUtil.crearCeldaCompacta("SALIDA ANTICIPADA", fontHeaderCompacto, colorPrincipal, 2, 1));
            encabezado
                .addCell(ReporteUtil.crearCeldaCompacta("T. ALIMENTACIÓN", fontHeaderCompacto, colorPrincipal, 1, 3));
            encabezado
                .addCell(
                    ReporteUtil.crearCeldaCompacta("TIEMPO PLANIFICADO", fontHeaderCompacto, colorPrincipal, 2, 1));
            encabezado
                .addCell(ReporteUtil.crearCeldaCompacta("TIEMPO LABORADO", fontHeaderCompacto, colorPrincipal, 2, 1));
            encabezado
                .addCell(ReporteUtil.crearCeldaCompacta("OBSERVACIONES", fontHeaderCompacto, colorPrincipal, 2, 1));

            for (int i = 0; i < 4; i++) {
              encabezado.addCell(ReporteUtil.crearCeldaCompacta("HORARIO", fontHeaderCompacto, colorPrincipal));
              encabezado.addCell(ReporteUtil.crearCeldaCompacta("TIMBRE", fontHeaderCompacto, colorSecundario));
              encabezado.addCell(ReporteUtil.crearCeldaCompacta("EST", fontHeaderCompacto, colorPrincipal));
            }

            encabezado.addCell(ReporteUtil.crearCeldaCompacta("ASIGNADO", fontHeaderCompacto, colorPrincipal));
            encabezado.addCell(ReporteUtil.crearCeldaCompacta("TOMADO", fontHeaderCompacto, colorPrincipal));
            encabezado.addCell(ReporteUtil.crearCeldaCompacta("EXCESO", fontHeaderCompacto, colorPrincipal));
            encabezado.setSpacingAfter(0f);
            document.add(encabezado);

            PdfPTable tablaData = new PdfPTable(22);
            tablaData.setWidthPercentage(100);
            tablaData.setWidths(WIDTHS_22);

            int contador = 1;
            if (emp.getTLaborado() != null) {
              for (RegistroAsistenciaDTO reg : emp.getTLaborado()) {
                Color fondo = (contador % 2 == 0) ? zebraColor : Color.WHITE;

                MarcaDTO entradaMarca = reg.getEntrada();
                MarcaAlimentacionDTO inicioMarca = reg.getInicioAlimentacion();
                MarcaAlimentacionDTO finMarca = reg.getFinAlimentacion();
                MarcaDTO salidaMarca = reg.getSalida();

                String entradaHorario = extraerHora(entradaMarca != null ? entradaMarca.getFecha_hora_horario() : null);
                String entradaTimbre = obtenerTextoTimbre(entradaMarca);
                String entradaEstado = obtenerTextoEstado(entradaMarca, reg.getOrigen(), reg.getControl());

                String inicioHorario = "EAS".equals(safe(reg.getTipo()))
                    ? extraerHora(inicioMarca != null ? inicioMarca.getFecha_hora_horario() : null)
                    : "";
                String inicioTimbre = "EAS".equals(safe(reg.getTipo()))
                    ? obtenerTextoTimbre(inicioMarca)
                    : "";
                String inicioEstado = "EAS".equals(safe(reg.getTipo()))
                    ? obtenerTextoEstado(inicioMarca, reg.getOrigen(), reg.getControl())
                    : "";

                String finHorario = "EAS".equals(safe(reg.getTipo()))
                    ? extraerHora(finMarca != null ? finMarca.getFecha_hora_horario() : null)
                    : "";
                String finTimbre = "EAS".equals(safe(reg.getTipo()))
                    ? obtenerTextoTimbre(finMarca)
                    : "";
                String finEstado = "EAS".equals(safe(reg.getTipo()))
                    ? obtenerTextoEstado(finMarca, reg.getOrigen(), reg.getControl())
                    : "";

                String salidaHorario = extraerHora(salidaMarca != null ? salidaMarca.getFecha_hora_horario() : null);
                String salidaTimbre = obtenerTextoTimbre(salidaMarca);
                String salidaEstado = obtenerTextoEstado(salidaMarca, reg.getOrigen(), reg.getControl());

                /////////////////
                boolean alimentacionEsPermiso = "P".equals(inicioEstado) || "P".equals(finEstado);
                boolean alimentacionEsJHE = "JHE".equalsIgnoreCase(finEstado);
                boolean atrasoEsJHE = "JHE".equalsIgnoreCase(entradaEstado);

                Double minutosAsignadosAlim = (inicioMarca != null && inicioMarca.getMinutos_alimentacion() != null)
                    ? inicioMarca.getMinutos_alimentacion()
                    : 0d;

                Double minLaborados = reg.getMinLaborados() != null ? reg.getMinLaborados() : 0d;

                Double minPlanificadosBase = reg.getMinPlanificados() != null ? reg.getMinPlanificados() : 0d;
                Double minPlanificados = minPlanificadosBase - minutosAsignadosAlim;

                if (minPlanificados < 0) {
                  minPlanificados = 0d;
                }

                Double minAtrasos = reg.getMinAtrasos() != null ? reg.getMinAtrasos() : 0d;
                Double minSalidas = reg.getMinSalidasAnticipadas() != null ? reg.getMinSalidasAnticipadas() : 0d;
                Double minAlimentacion = reg.getMinAlimentacion() != null ? reg.getMinAlimentacion() : 0d;
                Double minExcesoAlimentacion = calcularExcesoAlimentacion(minutosAsignadosAlim, minAlimentacion);

                //////

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(String.valueOf(contador), fontTextoCompacto, fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    ReporteUtil.formatearFechaConDia(entradaMarca != null ? entradaMarca.getFecha_horario() : null),
                    fontTextoCompacto, fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(entradaHorario, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(entradaTimbre, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    entradaEstado,
                    fontTextoCompacto,
                    getColorEstado(
                        entradaEstado,
                        fondo,
                        COLOR_FALTA_TIMBRE,
                        COLOR_PERMISO,
                        COLOR_VACACIONES,
                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(inicioHorario, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(inicioTimbre, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    inicioEstado,
                    fontTextoCompacto,
                    getColorEstado(
                        inicioEstado,
                        fondo,
                        COLOR_FALTA_TIMBRE,
                        COLOR_PERMISO,
                        COLOR_VACACIONES,
                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(finHorario, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(finTimbre, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    finEstado,
                    fontTextoCompacto,
                    getColorEstado(
                        finEstado,
                        fondo,
                        COLOR_FALTA_TIMBRE,
                        COLOR_PERMISO,
                        COLOR_VACACIONES,
                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(salidaHorario, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(salidaTimbre, fontTextoCompacto, fondo));
                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    salidaEstado,
                    fontTextoCompacto,
                    getColorEstado(
                        salidaEstado,
                        fondo,
                        COLOR_FALTA_TIMBRE,
                        COLOR_PERMISO,
                        COLOR_VACACIONES,
                        COLOR_JUSTIFICACION_HORAS_EXTRAS)));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minAtrasos),
                    fontTextoCompacto,
                    minAtrasos > 0 && !atrasoEsJHE ? COLOR_ATRASO : fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minSalidas),
                    fontTextoCompacto,
                    minSalidas > 0 ? COLOR_SALIDA_ANTICIPADA : fondo));

                boolean tieneExcesoAlimentacion = !alimentacionEsPermiso
                    && !alimentacionEsJHE
                    && minutosAsignadosAlim != null
                    && minAlimentacion > minutosAsignadosAlim;

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minutosAsignadosAlim),
                    fontTextoCompacto,
                    fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minAlimentacion),
                    fontTextoCompacto,
                    tieneExcesoAlimentacion ? COLOR_EXCESO_ALIMENTACION : fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minExcesoAlimentacion),
                    fontTextoCompacto,
                    tieneExcesoAlimentacion ? COLOR_EXCESO_ALIMENTACION : fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minPlanificados),
                    fontTextoCompacto,
                    fondo));

                tablaData.addCell(ReporteUtil.crearCeldaCompacta(
                    convertirMinutosATiempo(minLaborados),
                    fontTextoCompacto,
                    fondo));

                tablaData.addCell(ReporteUtil.crearCeldaObservacionConEstado(
                    safe(reg.getObservaciones()),
                    fondo));

                /*
                 * totalAtrasos += minAtrasos;
                 * totalSalidasAnticipadas += minSalidas;
                 * totalAlimentacionTomado += minAlimentacion;
                 * totalAlimentacionAsignado += minutosAsignadosAlim != null ?
                 * minutosAsignadosAlim : 0;
                 * totalExcesoAlimentacion += minExcesoAlimentacion != null ?
                 * minExcesoAlimentacion : 0;
                 * totalPlanificado += minPlanificados;
                 * totalLaborado += minLaborados;
                 */

                totalAtrasos += atrasoEsJHE ? 0d : minAtrasos;
                totalSalidasAnticipadas += minSalidas;
                totalAlimentacionTomado += tieneExcesoAlimentacion ? minAlimentacion : 0d;
                totalAlimentacionAsignado += minutosAsignadosAlim != null ? minutosAsignadosAlim : 0;
                totalExcesoAlimentacion += tieneExcesoAlimentacion && minExcesoAlimentacion != null
                    ? minExcesoAlimentacion
                    : 0d;
                totalPlanificado += minPlanificados;
                totalLaborado += minLaborados;

                contador++;
              }
            }

            for (int i = 0; i < 13; i++) {
              PdfPCell celdaVacia = ReporteUtil.crearCeldaCompacta("", fontTextoCompacto, Color.WHITE);
              celdaVacia.setBorder(Rectangle.NO_BORDER);
              tablaData.addCell(celdaVacia);
            }

            tablaData.addCell(ReporteUtil.crearCeldaCompacta("TOTAL", fontTextoCompacto, Color.WHITE));
            tablaData.addCell(
                ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalAtrasos), fontTextoCompacto, Color.WHITE));
            tablaData.addCell(ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalSalidasAnticipadas),
                fontTextoCompacto, Color.WHITE));
            tablaData.addCell(ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalAlimentacionAsignado),
                fontTextoCompacto, Color.WHITE));
            tablaData.addCell(ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalAlimentacionTomado),
                fontTextoCompacto, Color.WHITE));
            tablaData.addCell(ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalExcesoAlimentacion),
                fontTextoCompacto, Color.WHITE));
            tablaData.addCell(ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalPlanificado),
                fontTextoCompacto, Color.WHITE));
            tablaData.addCell(
                ReporteUtil.crearCeldaCompacta(convertirMinutosATiempo(totalLaborado), fontTextoCompacto, Color.WHITE));
            tablaData.addCell(ReporteUtil.crearCeldaCompacta("", fontTextoCompacto, Color.WHITE));

            tablaData.setSpacingAfter(10f);
            document.add(tablaData);
          }
        }
      }

      document.close();
      return baos.toByteArray();

    } catch (IllegalArgumentException e) {
      System.out.println("Error en solicitud: " + e.getMessage());
      throw e;
    } catch (Exception e) {
      System.out.println("Error al generar ResumenAsistencia.pdf: " + e.getMessage());
      throw new ReportBuildException("No se pudo generar ResumenAsistencia.pdf", e);
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

  public byte[] generarReporteResumenAsistenciaXLSX(ReporteAsistenciaRequest request) {
    final String NOMBRE_HOJA = "Resumen_asistencia";
    final int FILA_ENCABEZADO = 5;

    final int MERGE_FIL_INI = 0, MERGE_FIL_FIN = 4;
    final int MERGE_COL_INI = 1, MERGE_COL_FIN = 28;

    final String[] HEADERS = {
        "ITEM",
        "IDENTIFICACIÓN",
        "CÓDIGO",
        "APELLIDO NOMBRE",
        "CIUDAD",
        "SUCURSAL",
        "RÉGIMEN",
        "DEPARTAMENTO",
        "CARGO",
        "FECHA",

        "HORARIO ENTRADA",
        "TIMBRE ENTRADA",
        "EST ENTRADA",

        "HORARIO INICIO ALIMENTACIÓN",
        "TIMBRE INICIO ALIMENTACIÓN",
        "EST INICIO ALIMENTACIÓN",

        "HORARIO FIN ALIMENTACIÓN",
        "TIMBRE FIN ALIMENTACIÓN",
        "EST FIN ALIMENTACIÓN",

        "HORARIO SALIDA",
        "TIMBRE SALIDA",
        "EST SALIDA",

        "ATRASO",
        "SALIDA ANTICIPADA",
        "TIEMPO ALIMENTACIÓN ASIGNADO",
        "TIEMPO ALIMENTACIÓN TOMADO HH:MM:SS",
        "TIEMPO ALIMENTACIÓN EXCESO HH:MM:SS",
        "TIEMPO PLANIFICADO HH:MM:SS",
        "TIEMPO LABORADO HH:MM:SS",
        "OBSERVACIONES"
    };

    final int[] ANCHOS = {
        10, 18, 12, 28, 16, 16, 18, 20, 18, 16,
        14, 14, 12,
        18, 18, 12,
        18, 18, 12,
        14, 14, 12,
        14, 18, 18, 18, 18, 18, 18,
        45
    };

    try (XSSFWorkbook libro = new XSSFWorkbook();
        ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

      XSSFSheet hoja = libro.createSheet(NOMBRE_HOJA);
      hoja.createFreezePane(0, FILA_ENCABEZADO + 1);

      byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
      if (logo != null && logo.length > 0) {
        UtilExcel.insertarLogoEstandar(libro, hoja, logo);
      }

      for (int row = MERGE_FIL_INI; row <= MERGE_FIL_FIN; row++) {
        UtilExcel.combinarCeldas(hoja, row, row, MERGE_COL_INI, MERGE_COL_FIN);
      }

      CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
      UtilExcel.establecerTexto(hoja, 0, 1, UtilExcel.aMayusculasSeguras(request.getEmpresa()), estiloTitulo);
      UtilExcel.establecerTexto(hoja, 1, 1, "RESUMEN DE ASISTENCIA", estiloTitulo);
      String periodo = "PERIODO DEL REPORTE: " + safe(request.getFechaInicio()) + " AL " + safe(request.getFechaFin());
      UtilExcel.establecerTexto(hoja, 2, 1, periodo, estiloTitulo);

      Row filaHeader = UtilExcel.asegurarFila(hoja, FILA_ENCABEZADO);
      for (int c = 0; c < HEADERS.length; c++) {
        UtilExcel.establecerTexto(filaHeader, c, HEADERS[c], null);
      }

      CellStyle estiloHeader = ConfiguracionExcel.crearEstiloEncabezadoTabla(libro);
      UtilExcel.aplicarEstiloAFila(filaHeader, HEADERS.length, estiloHeader);
      UtilExcel.establecerAnchosColumnas(hoja, ANCHOS);
      hoja.getRow(FILA_ENCABEZADO).setHeightInPoints(18f);

      int filaDatosIni = FILA_ENCABEZADO + 1;
      int filaAct = filaDatosIni;
      int item = 1;

      if (request.getGrupos() != null) {
        for (GrupoAsistenciaDTO grupo : request.getGrupos()) {
          if (grupo.getEmpleados() == null) {
            continue;
          }

          for (EmpleadoAsistenciaDTO usu : grupo.getEmpleados()) {
            String apenom = (safe(usu.getApellido()) + " " + safe(usu.getNombre())).trim();
            if (usu.getTLaborado() == null) {
              continue;
            }

            for (RegistroAsistenciaDTO t : usu.getTLaborado()) {
              Row r = UtilExcel.asegurarFila(hoja, filaAct++);
              int col = 0;

              MarcaDTO entrada = t.getEntrada();
              MarcaAlimentacionDTO inicio = t.getInicioAlimentacion();
              MarcaAlimentacionDTO fin = t.getFinAlimentacion();
              MarcaDTO salida = t.getSalida();

              String entradaHorario = extraerHora(entrada != null ? entrada.getFecha_hora_horario() : null);
              String entradaTimbre = obtenerTextoTimbre(entrada);
              String entradaEstado = obtenerTextoEstado(entrada, safe(t.getOrigen()), t.getControl());

              String inicioHorario = "EAS".equals(safe(t.getTipo()))
                  ? extraerHora(inicio != null ? inicio.getFecha_hora_horario() : null)
                  : "";
              String inicioTimbre = "EAS".equals(safe(t.getTipo()))
                  ? obtenerTextoTimbre(inicio)
                  : "";
              String inicioEstado = "EAS".equals(safe(t.getTipo()))
                  ? obtenerTextoEstado(inicio, safe(t.getOrigen()), t.getControl())
                  : "";

              String finHorario = "EAS".equals(safe(t.getTipo()))
                  ? extraerHora(fin != null ? fin.getFecha_hora_horario() : null)
                  : "";
              String finTimbre = "EAS".equals(safe(t.getTipo()))
                  ? obtenerTextoTimbre(fin)
                  : "";
              String finEstado = "EAS".equals(safe(t.getTipo()))
                  ? obtenerTextoEstado(fin, safe(t.getOrigen()), t.getControl())
                  : "";

              String salidaHorario = extraerHora(salida != null ? salida.getFecha_hora_horario() : null);
              String salidaTimbre = obtenerTextoTimbre(salida);
              String salidaEstado = obtenerTextoEstado(salida, safe(t.getOrigen()), t.getControl());

              Double minsLaborados = t.getMinLaborados() != null ? t.getMinLaborados() : 0d;

              Double asignMin = ("EAS".equals(safe(t.getTipo()))
                  && inicio != null
                  && inicio.getMinutos_alimentacion() != null)
                      ? inicio.getMinutos_alimentacion()
                      : 0d;

              Double minsPlanificadosBase = t.getMinPlanificados() != null ? t.getMinPlanificados() : 0d;
              Double minsPlanificados = minsPlanificadosBase - asignMin;

              if (minsPlanificados < 0) {
                minsPlanificados = 0d;
              }

              Double minTomadoAlimentacion = t.getMinAlimentacion() != null ? t.getMinAlimentacion() : 0d;
              Double minExcesoAlimentacion = calcularExcesoAlimentacion(asignMin, minTomadoAlimentacion);
              UtilExcel.establecerValor(r, col++, item++, null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getIdentificacion()), null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getCodigo()), null);
              UtilExcel.establecerTexto(r, col++, apenom, null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getCiudad()), null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getSucursal()), null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getRegimen()), null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getDepartamento()), null);
              UtilExcel.establecerTexto(r, col++, safe(usu.getCargo()), null);
              UtilExcel.establecerTexto(r, col++, safe(entrada != null ? entrada.getFecha_horario() : null), null);

              UtilExcel.establecerTexto(r, col++, entradaHorario, null);
              UtilExcel.establecerTexto(r, col++, entradaTimbre, null);
              UtilExcel.establecerTexto(r, col++, entradaEstado, null);

              UtilExcel.establecerTexto(r, col++, inicioHorario, null);
              UtilExcel.establecerTexto(r, col++, inicioTimbre, null);
              UtilExcel.establecerTexto(r, col++, inicioEstado, null);

              UtilExcel.establecerTexto(r, col++, finHorario, null);
              UtilExcel.establecerTexto(r, col++, finTimbre, null);
              UtilExcel.establecerTexto(r, col++, finEstado, null);

              UtilExcel.establecerTexto(r, col++, salidaHorario, null);
              UtilExcel.establecerTexto(r, col++, salidaTimbre, null);
              UtilExcel.establecerTexto(r, col++, salidaEstado, null);

              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(t.getMinAtrasos()), null);
              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(t.getMinSalidasAnticipadas()), null);
              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(asignMin), null);
              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(minTomadoAlimentacion), null);
              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(minExcesoAlimentacion), null);
              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(minsPlanificados), null);
              UtilExcel.establecerTexto(r, col++, convertirMinutosATiempo(minsLaborados), null);
              UtilExcel.establecerTexto(r, col++, safe(t.getObservaciones()), null);
            }
          }
        }
      }

      int ultimaFila = (filaAct == filaDatosIni) ? FILA_ENCABEZADO : (filaAct - 1);

      CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
      CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

      UtilExcel.aplicarEstiloARegion(hoja, FILA_ENCABEZADO, FILA_ENCABEZADO, 0, HEADERS.length - 1,
          estiloCentroBorde, true);

      if (ultimaFila >= filaDatosIni) {
        UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 0, 0, estiloCentroBorde, true);
        UtilExcel.aplicarEstiloARegion(hoja, filaDatosIni, ultimaFila, 1, HEADERS.length - 1, estiloIzqBorde, true);
      }

      if (ultimaFila >= filaDatosIni) {
        boolean[] filtros = new boolean[HEADERS.length];
        for (int i = 0; i < filtros.length; i++) {
          filtros[i] = true;
        }
        filtros[0] = false;

        UtilExcel.crearTablaEstilizada(
            hoja,
            "ResumenGeneralReporteTabla",
            FILA_ENCABEZADO, 0,
            ultimaFila, HEADERS.length - 1,
            true,
            filtros);
      }

      libro.write(baos);
      return baos.toByteArray();

    } catch (IllegalArgumentException e) {
      throw e;
    } catch (Exception e) {
      throw new ReportBuildException("No se pudo generar ResumenAsistencia.xlsx", e);
    }
  }

  private String safe(Object v) {
    if (v == null) {
      return "";
    }
    String s = String.valueOf(v).trim();
    return "null".equalsIgnoreCase(s) ? "" : s;
  }

  private String convertirMinutosATiempo(Double minutos) {
    if (minutos == null || minutos <= 0) {
      return "00:00:00";
    }
    int totalSegundos = (int) Math.round(minutos * 60);
    int horas = totalSegundos / 3600;
    int mins = (totalSegundos % 3600) / 60;
    int segundos = totalSegundos % 60;
    return String.format("%02d:%02d:%02d", horas, mins, segundos);
  }

  private String extraerHora(String fechaHora) {
    if (fechaHora == null || !fechaHora.contains(" ")) {
      return "";
    }
    String[] partes = fechaHora.split(" ");
    return partes.length > 1 ? partes[1] : "";
  }

  private String obtenerTextoTimbre(MarcaDTO marca) {
    if (marca == null) {
      return "";
    }
    if (marca.getFecha_hora_timbre() != null && !marca.getFecha_hora_timbre().trim().isEmpty()) {
      return extraerHora(marca.getFecha_hora_timbre());
    }
    return "";
  }

  private String obtenerTextoEstado(MarcaDTO marca, String origen, Boolean control) {
    if (marca == null) {
      return "";
    }

    String estadoTimbre = safe(marca.getEstado_timbre());

    if (!estadoTimbre.isEmpty()) {
      return estadoTimbre;
    }

    if ("L".equals(origen) || "FD".equals(origen) || "DHA".equals(origen)) {
      return origen;
    }

    return (control != null && control) ? "FT" : "SCA";
  }

  private Color getColorEstado(
      String valor,
      Color porDefecto,
      Color colorFT,
      Color colorPermiso,
      Color colorVacaciones,
      Color colorJHE) {

    String estado = safe(valor);

    if ("FT".equalsIgnoreCase(estado)) {
      return colorFT;
    }
    if ("P".equalsIgnoreCase(estado)) {
      return colorPermiso;
    }
    if ("V".equalsIgnoreCase(estado)) {
      return colorVacaciones;
    }
    if ("JHE".equalsIgnoreCase(estado)) {
      return colorJHE;
    }
    return porDefecto;
  }

  private Double calcularExcesoAlimentacion(Double minutosAsignados, Double minutosTomados) {
    double asignado = minutosAsignados != null ? minutosAsignados : 0d;
    double tomado = minutosTomados != null ? minutosTomados : 0d;

    if (tomado > asignado) {
      return tomado - asignado;
    }

    return 0d;
  }

}