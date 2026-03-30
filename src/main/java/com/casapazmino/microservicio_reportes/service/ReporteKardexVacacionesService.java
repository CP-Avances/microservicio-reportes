package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionExcel;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.UtilExcel;

import org.openpdf.text.Chunk;
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
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

@Service
public class ReporteKardexVacacionesService {

  // =========================================================================================
  // PDF
  // =========================================================================================
  public byte[] generarReporteKardexVacacionesPDF(ReporteKardexVacacionesRequest request) {

    // anchos
    final float[] W_EMP_INFO = { 4f, 4f, 2f };
    final float[] W_PERIODO_FILA = { 2.4f, 2.4f, 2.4f, 2.0f, 1.8f };
    final float[] W_MOV = { 3.5f, 1.7f, 1.2f, 1.7f, 1.2f, 0.8f, 0.8f, 0.8f, 0.8f, 0.8f, 0.8f };
    final float[] W_PEND = { 3.5f, 1.7f, 1.2f, 1.7f, 1.2f, 0.8f, 0.8f, 0.8f, 1.4f };

    Document document = null;
    PdfWriter writer = null;
    ByteArrayOutputStream baos = null;

    try {
      baos = new ByteArrayOutputStream();
      document = new Document(PageSize.A4, 40, 40, 30, 50);
      writer = PdfWriter.getInstance(document, baos);

      writer.setPageEvent(new ConfiguracionPaginaPDF(
          safe(request.getUsuario()),
          safe(request.getFraseMarcaAgua()),
          safe(request.getColorPrincipal())));

      document.open();

      // Logo
      Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
      if (logo != null)
        document.add(logo);

      // Títulos
      document.add(ReporteUtil.crearTituloEmpresa(safe(request.getEmpresa())));
      document.add(ReporteUtil.crearTituloReporte("KARDEX VACACIONES"));
      document.add(ReporteUtil.crearTituloPeriodo("FECHA DE CORTE: " + safe(request.getFechaCorte())));

      Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
      Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
      Color zebra = ReporteUtil.colorZebraClaro();

      List<KardexEmpleadoDTO> empleadosIn = request.getEmpleados();
      if (empleadosIn == null || empleadosIn.isEmpty()) {
        Paragraph p = new Paragraph("Sin datos para mostrar", ReporteUtil.fuenteTexto());
        p.setAlignment(Element.ALIGN_CENTER);
        p.setSpacingBefore(20f);
        document.add(p);
        document.close();
        return baos.toByteArray();
      }

      // ==========================================================
      // ✅ ORDENAR Y AGRUPAR SEGÚN CRITERIO (NO MEZCLAR)
      // ==========================================================
      final String criterio = safe(request.getCriterio()).toLowerCase();

      List<KardexEmpleadoDTO> listaOrdenada = empleadosIn.stream()
          .filter(Objects::nonNull)
          .sorted(Comparator
              .comparing((KardexEmpleadoDTO e) -> safe(grupoNombre(criterio, e)))
              .thenComparing(e -> safe(e.getApellido()))
              .thenComparing(e -> safe(e.getNombre())))
          .collect(Collectors.toList());

      LinkedHashMap<String, List<KardexEmpleadoDTO>> grupos = listaOrdenada.stream()
          .collect(Collectors.groupingBy(
              e -> safe(grupoNombre(criterio, e)),
              LinkedHashMap::new,
              Collectors.toList()));

      // =====================================================================
      // EMPLEADOS (agrupados)
      // =====================================================================
      for (Map.Entry<String, List<KardexEmpleadoDTO>> entry : grupos.entrySet()) {
        List<KardexEmpleadoDTO> empsGrupo = entry.getValue();
        if (empsGrupo == null || empsGrupo.isEmpty())
          continue;

        // ---- empleados dentro del grupo ----
        for (KardexEmpleadoDTO emp : empsGrupo) {
          if (emp == null)
            continue;

          // ✅ FRANJA VERDE POR EMPLEADO (según criterio) (EXCEPTO criterio 'e')
          String franja = textoFranjaEmpleado(criterio, emp);
          if (!franja.isBlank()) {
            PdfPTable tFr = new PdfPTable(1);
            tFr.setWidthPercentage(100);

            PdfPCell cFr = new PdfPCell(new Phrase(franja, ReporteUtil.fuenteEncabezado()));
            cFr.setPadding(6f);
            cFr.setBorder(Rectangle.BOX);
            cFr.setBorderColor(new Color(200, 200, 200));
            cFr.setBackgroundColor(new Color(230, 255, 230)); // verde suave

            tFr.addCell(cFr);
            tFr.setSpacingBefore(10f);
            tFr.setSpacingAfter(6f);
            document.add(tFr);
          }

          // ---- Cabecera empleado (SOLO: CIUDAD / CC / COD + EMPLEADO) ----
          PdfPTable infoEmp = new PdfPTable(3);
          infoEmp.setWidthPercentage(100);
          infoEmp.setWidths(W_EMP_INFO);

          infoEmp.addCell(celdaInfoMixtaLocal("CIUDAD:", safe(emp.getCiudad()), zebra));
          infoEmp.addCell(celdaInfoMixtaLocal("C.C.:", safe(emp.getIdentificacion()), zebra));
          infoEmp.addCell(celdaInfoMixtaLocal("COD:", safe(emp.getCodigo()), zebra));
          String nombreCompleto = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();

          PdfPCell cNombre = celdaInfoMixtaLocal("EMPLEADO:", nombreCompleto, zebra);
          cNombre.setColspan(3);
          cNombre.setPaddingBottom(4f);
          infoEmp.addCell(cNombre);

          PdfPTable contInfoEmp = new PdfPTable(1);
          contInfoEmp.setWidthPercentage(100);
          PdfPCell wrap = new PdfPCell(infoEmp);
          wrap.setBorder(Rectangle.BOX);
          wrap.setPadding(0);
          contInfoEmp.addCell(wrap);
          contInfoEmp.setSpacingBefore(4f);
          contInfoEmp.setSpacingAfter(8f);
          noSpace(contInfoEmp);

          for (KardexPeriodoDTO per : emp.getPeriodos()) {
            if (per == null)
              continue;

            // ==========================
            // FILA ENCERRADA (periodo)
            // ==========================
            PdfPTable filaPeriodo = new PdfPTable(5);
            filaPeriodo.setWidthPercentage(100);
            filaPeriodo.setWidths(W_PERIODO_FILA);

            filaPeriodo.addCell(ReporteUtil.crearCelda("F.INICIO " + safe(per.getF_ingreso()),
                ReporteUtil.fuenteTexto(), Color.WHITE));
            filaPeriodo.addCell(ReporteUtil.crearCelda("F.FIN " + safe(per.getF_salida()),
                ReporteUtil.fuenteTexto(), Color.WHITE));
            filaPeriodo.addCell(ReporteUtil.crearCelda("F.CARGA " + safe(per.getF_carga()),
                ReporteUtil.fuenteTexto(), Color.WHITE));
            filaPeriodo.addCell(ReporteUtil.crearCelda("DÍAS ACUM " + fmtDec2(per.getDias_acum_decimal()),
                ReporteUtil.fuenteTexto(), Color.WHITE));
            String est = "ACTIVO".equalsIgnoreCase(safe(per.getEstado_periodo())) ? "Activo" : "Inactivo";
            filaPeriodo.addCell(
                ReporteUtil.crearCelda("ESTADO " + est, ReporteUtil.fuenteTexto(), Color.WHITE));

            // --- contenedor del periodo SIN aire ---
            noSpace(filaPeriodo);

            PdfPTable contFila = new PdfPTable(1);
            contFila.setWidthPercentage(100);
            noSpace(contFila);

            PdfPCell filaWrap = new PdfPCell(filaPeriodo);
            filaWrap.setPadding(0f); // <-- sin aire
            filaWrap.setBorder(Rectangle.BOX); // si quieres marco
            filaWrap.setBorderColor(new Color(200, 200, 200));
            filaWrap.setBorderWidth(0.8f);

            contFila.addCell(filaWrap);

            // ==========================
            // TABLA MOVIMIENTOS
            // ==========================
            PdfPTable mov = new PdfPTable(11);
            mov.setWidthPercentage(100);
            mov.setWidths(W_MOV);

            mov.addCell(hCell("Detalle", colorPrincipal, 1, 2));
            mov.addCell(hCell("Desde", colorPrincipal, 2, 1));
            mov.addCell(hCell("Hasta", colorPrincipal, 2, 1));
            mov.addCell(hCell("Descuento", colorPrincipal, 3, 1));
            mov.addCell(hCell("Saldo", colorPrincipal, 3, 1));

            mov.addCell(hCell("Fecha", colorPrincipal));
            mov.addCell(hCell("Hora", colorPrincipal));
            mov.addCell(hCell("Fecha", colorPrincipal));
            mov.addCell(hCell("Hora", colorPrincipal));
            mov.addCell(hCell("Días", colorPrincipal));
            mov.addCell(hCell("Hor", colorPrincipal));
            mov.addCell(hCell("Min", colorPrincipal));
            mov.addCell(hCell("Días", colorPrincipal));
            mov.addCell(hCell("Hor", colorPrincipal));
            mov.addCell(hCell("Min", colorPrincipal));

            List<KardexMovimientoDTO> movimientos = per.getMovimientos();
            if (movimientos == null || movimientos.isEmpty()) {
              PdfPCell sin = new PdfPCell(new Phrase(
                  "Sin movimientos autorizados hasta la fecha de corte", ReporteUtil.fuenteTablaData()));
              sin.setColspan(11);
              sin.setHorizontalAlignment(Element.ALIGN_CENTER);
              sin.setPadding(6f);
              mov.addCell(sin);
            } else {
              int i = 1;
              for (KardexMovimientoDTO m : movimientos) {
                Color fondo = (i % 2 == 0) ? zebra : Color.WHITE;

                mov.addCell(cellLeft(safe(m.getDetalle()), fondo));
                mov.addCell(cellCenter(safe(m.getFecha_inicio()), fondo));
                mov.addCell(cellCenter(safe(m.getHora_inicio()), fondo));
                mov.addCell(cellCenter(safe(m.getFecha_final()), fondo));
                mov.addCell(cellCenter(safe(m.getHora_fin()), fondo));

                DHMDTO d = m.getDescuento_dhm();
                mov.addCell(cellCenter(String.valueOf(int0(d != null ? d.getDias() : 0)), fondo));
                mov.addCell(cellCenter(String.valueOf(int0(d != null ? d.getHoras() : 0)), fondo));
                mov.addCell(cellCenter(String.valueOf(int0(d != null ? d.getMinutos() : 0)), fondo));

                DHMDTO s = m.getSaldo_dhm();
                mov.addCell(cellCenter(String.valueOf(int0(s != null ? s.getDias() : 0)), fondo));
                mov.addCell(cellCenter(String.valueOf(int0(s != null ? s.getHoras() : 0)), fondo));
                mov.addCell(cellCenter(String.valueOf(int0(s != null ? s.getMinutos() : 0)), fondo));
                i++;
              }
            }
            mov.setSpacingAfter(8f);
            noSpace(mov);

            // ==========================
            // ANTIGÜEDAD (izq) + PROPORCIONAL (bajo DESCUENTO) + LIQUIDACIÓN (bajo SALDO)
            // ✅ Alineado con columnas de la tabla MOV (W_MOV)
            // ==========================
            PdfPTable bloqueAlineado = new PdfPTable(11);
            bloqueAlineado.setWidthPercentage(100);
            bloqueAlineado.setWidths(W_MOV);
            bloqueAlineado.setSpacingAfter(8f);

            // ---- Colores estilo ejemplo (azul claro) ----
            Color azulHeader = new Color(183, 208, 230);
            Color azulSub = new Color(210, 225, 240);

            // ====== 1) ANTIGÜEDAD (colspan 3, lado izquierdo) ======
            PdfPTable tAnt = new PdfPTable(1);
            tAnt.setWidthPercentage(100);
            tAnt.addCell(hCell("Días Antigüedad", azulHeader));
            tAnt.addCell(cellCenter(String.valueOf(int0(per.getDias_antiguedad())), Color.WHITE));

            // ====== 1) ANTIGÜEDAD (SOLO col Detalle: colspan 1) ======
            PdfPCell cAnt = new PdfPCell(tAnt);
            cAnt.setColspan(1); // SOLO Detalle
            cAnt.setPadding(0);
            cAnt.setBorder(Rectangle.BOX);
            cAnt.setBorderColor(new Color(80, 80, 80));
            cAnt.setBorderWidth(0.8f);
            bloqueAlineado.addCell(cAnt);

            // separadores para Desde(Fecha/Hora) y Hasta(Fecha/Hora)
            for (int k = 0; k < 4; k++) {
              PdfPCell sep = new PdfPCell(new Phrase(""));
              sep.setColspan(1);
              sep.setBorder(Rectangle.NO_BORDER);
              sep.setFixedHeight(1f);
              bloqueAlineado.addCell(sep);
            }

            // ====== 2) PROPORCIONAL (colspan 3, EXACTO bajo DESCUENTO) ======
            PdfPTable tProp = new PdfPTable(3);
            tProp.setWidthPercentage(100);
            tProp.setWidths(new float[] { 1f, 1f, 1f });

            tProp.addCell(hCell("Proporcional", azulHeader, 3, 1));
            tProp.addCell(hCell("Días", azulSub));
            tProp.addCell(hCell("Hor", azulSub));
            tProp.addCell(hCell("Min", azulSub));

            DHMDTO pd = per.getProporcional_dhm();
            tProp.addCell(cellCenter(String.valueOf(int0(pd != null ? pd.getDias() : 0)), Color.WHITE));
            tProp.addCell(cellCenter(String.valueOf(int0(pd != null ? pd.getHoras() : 0)), Color.WHITE));
            tProp.addCell(cellCenter(String.valueOf(int0(pd != null ? pd.getMinutos() : 0)), Color.WHITE));

            PdfPCell cProp = new PdfPCell(tProp);
            cProp.setColspan(3);
            cProp.setPadding(0);
            cProp.setBorder(Rectangle.BOX);
            cProp.setBorderColor(new Color(80, 80, 80));
            cProp.setBorderWidth(0.8f);
            bloqueAlineado.addCell(cProp);

            // ====== 3) LIQUIDACIÓN (colspan 3, EXACTO bajo SALDO) ======
            PdfPTable tLiq = new PdfPTable(3);
            tLiq.setWidthPercentage(100);
            tLiq.setWidths(new float[] { 1f, 1f, 1f });

            tLiq.addCell(hCell("Liquidación", azulHeader, 3, 1));
            tLiq.addCell(hCell("Días", azulSub));
            tLiq.addCell(hCell("Hor", azulSub));
            tLiq.addCell(hCell("Min", azulSub));

            DHMDTO ld = per.getLiquidacion_dhm();
            tLiq.addCell(cellCenter(String.valueOf(int0(ld != null ? ld.getDias() : 0)), Color.WHITE));
            tLiq.addCell(cellCenter(String.valueOf(int0(ld != null ? ld.getHoras() : 0)), Color.WHITE));
            tLiq.addCell(cellCenter(String.valueOf(int0(ld != null ? ld.getMinutos() : 0)), Color.WHITE));

            PdfPCell cLiq = new PdfPCell(tLiq);
            cLiq.setColspan(3);
            cLiq.setPadding(0);
            cLiq.setBorder(Rectangle.BOX);
            cLiq.setBorderColor(new Color(80, 80, 80));
            cLiq.setBorderWidth(0.8f);
            bloqueAlineado.addCell(cLiq);

            noSpace(bloqueAlineado);

            // ==========================
            // ✅ CONTENEDOR ÚNICO (pegado, sin espacios)
            // ==========================
            PdfPTable card = new PdfPTable(1);
            card.setWidthPercentage(100);
            noSpace(card);

            // agrega en orden, sin espacios
            card.addCell(wrapNoSpace(contInfoEmp));
            card.addCell(wrapNoSpace(contFila));
            card.addCell(wrapNoSpace(mov));
            card.addCell(wrapNoSpace(bloqueAlineado));

            // separación solo al final del bloque principal
            card.setSpacingAfter(6f);
            document.add(card);

            // ===== Distribución de días FUERA del card =====
            PdfPTable tituloDistrib = construirTituloDistribucionPDF();
            document.add(tituloDistrib);

            PdfPTable distrib = construirTablaDistribucionPDF(per, azulHeader, azulSub);
            document.add(distrib);

            // ==========================
            // PENDIENTES
            // ==========================
            List<KardexPendienteDTO> pendientes = per.getPendientes();
            if (pendientes != null && !pendientes.isEmpty()) {

              PdfPTable tituloPend = new PdfPTable(1);
              tituloPend.setWidthPercentage(100);
              PdfPCell tp = new PdfPCell(
                  new Phrase("Solicitudes de permisos y vacaciones (Pendientes, Pre-Autorizadas)",
                      ReporteUtil.fuenteEncabezado()));
              tp.setBackgroundColor(new Color(238, 243, 255));
              tp.setBorderColor(new Color(207, 216, 255));
              tp.setPadding(6f);
              tituloPend.addCell(tp);
              tituloPend.setSpacingAfter(4f);
              document.add(tituloPend);

              PdfPTable pend = new PdfPTable(9);
              pend.setWidthPercentage(100);
              pend.setWidths(W_PEND);

              pend.addCell(hCell("Detalle", colorPrincipal, 1, 2));
              pend.addCell(hCell("Desde", colorPrincipal, 2, 1));
              pend.addCell(hCell("Hasta", colorPrincipal, 2, 1));
              pend.addCell(hCell("Tiempo", colorPrincipal, 3, 1));
              pend.addCell(hCell("Estado", colorPrincipal, 1, 2));

              pend.addCell(hCell("Fecha", colorPrincipal));
              pend.addCell(hCell("Hora", colorPrincipal));
              pend.addCell(hCell("Fecha", colorPrincipal));
              pend.addCell(hCell("Hora", colorPrincipal));
              pend.addCell(hCell("Días", colorPrincipal));
              pend.addCell(hCell("Hor", colorPrincipal));
              pend.addCell(hCell("Min", colorPrincipal));

              int i = 1;
              double totalDec = 0.0;

              for (KardexPendienteDTO x : pendientes) {
                Color fondo = (i % 2 == 0) ? zebra : Color.WHITE;

                DHMDTO t = pendienteDHM(emp, x);
                totalDec += toDecimalDias(x.getDias(), x.getMinutos_totales(),
                    int0(emp.getMin_por_dia()));

                  pend.addCell(cellLeft(safe(x.getDetalle()), fondo));
                  pend.addCell(cellCenter(safe(x.getFecha_inicio()), fondo));
                  pend.addCell(cellCenter(safe(x.getHora_inicio()), fondo));
                  pend.addCell(cellCenter(safe(x.getFecha_final()), fondo));
                  pend.addCell(cellCenter(safe(x.getHora_fin()), fondo));

                  pend.addCell(cellCenter(String.valueOf(int0(t.getDias())), fondo));
                  pend.addCell(cellCenter(String.valueOf(int0(t.getHoras())), fondo));
                  pend.addCell(cellCenter(String.valueOf(int0(t.getMinutos())), fondo));

                  pend.addCell(cellCenter(safe(x.getEstado_texto()), fondo));
                i++;
              }

              DHMDTO total = decimalToDHM(totalDec, int0(emp.getMin_por_dia()));
              Color fondoTotal = new Color(183, 208, 230);

              // Detalle, Desde Fecha, Desde Hora, Hasta Fecha, Hasta Hora
              for (int k = 0; k < 5; k++) {
                PdfPCell c = new PdfPCell(new Phrase(""));
                c.setBackgroundColor(fondoTotal);
                c.setBorder(Rectangle.BOX);
                c.setBorderColor(new Color(80, 80, 80));
                c.setBorderWidth(0.6f);
                c.setPadding(6f);
                pend.addCell(c);
              }

              // Tiempo
              pend.addCell(cellCenter(String.valueOf(int0(total.getDias())), fondoTotal));
              pend.addCell(cellCenter(String.valueOf(int0(total.getHoras())), fondoTotal));
              pend.addCell(cellCenter(String.valueOf(int0(total.getMinutos())), fondoTotal));

              // Estado con texto
              PdfPCell cEstadoTotal = new PdfPCell(new Phrase(
                  "Total estimado:",
                  ReporteUtil.fuenteEncabezadoTablaData()));
              cEstadoTotal.setHorizontalAlignment(Element.ALIGN_CENTER);
              cEstadoTotal.setVerticalAlignment(Element.ALIGN_MIDDLE);
              cEstadoTotal.setBackgroundColor(fondoTotal);
              cEstadoTotal.setPadding(6f);
              cEstadoTotal.setBorder(Rectangle.BOX);
              cEstadoTotal.setBorderColor(new Color(80, 80, 80));
              cEstadoTotal.setBorderWidth(0.6f);
              pend.addCell(cEstadoTotal);

              pend.setSpacingAfter(10f);
              document.add(pend);
            }
          }

        }
      }

      document.close();
      return baos.toByteArray();

    } catch (IllegalArgumentException e) {
      throw e;
    } catch (Exception e) {
      throw new ReportBuildException("No se pudo generar KardexVacaciones.pdf", e);
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
  public byte[] generarReporteKardexVacacionesExcel(ReporteKardexVacacionesRequest request) {

    try (XSSFWorkbook libro = new XSSFWorkbook();
        ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

      // Estilos base
      CellStyle estiloTitulo = ConfiguracionExcel.crearEstiloTitulo(libro);
      CellStyle estiloCentroBorde = ConfiguracionExcel.crearEstiloCentroConBorde(libro);
      CellStyle estiloIzqBorde = ConfiguracionExcel.crearEstiloIzquierdaConBorde(libro);

      final String criterio = safe(request.getCriterio()).toLowerCase();
      List<KardexEmpleadoDTO> empleados = (request.getEmpleados() == null) ? new ArrayList<>()
          : request.getEmpleados();

      // Ordenar + agrupar (para no mezclar)
      List<KardexEmpleadoDTO> listaOrdenada = empleados.stream()
          .filter(Objects::nonNull)
          .sorted(Comparator
              .comparing((KardexEmpleadoDTO e) -> safe(grupoNombre(criterio, e)))
              .thenComparing(e -> safe(e.getApellido()))
              .thenComparing(e -> safe(e.getNombre())))
          .collect(Collectors.toList());

      LinkedHashMap<String, List<KardexEmpleadoDTO>> grupos = listaOrdenada.stream()
          .collect(Collectors.groupingBy(
              e -> safe(grupoNombre(criterio, e)),
              LinkedHashMap::new,
              Collectors.toList()));

      // =========================
      // ÚNICA HOJA: REPORTE (FORMATO PDF)
      // =========================
      final String HOJA_REP = "Kardex_Reporte";
      XSSFSheet shRep = libro.createSheet(HOJA_REP);

      byte[] logo = UtilExcel.decodificarImagenBase64(request.getLogoBase64());
      if (logo != null && logo.length > 0) {
        insertarLogoKardexA1A4(libro, shRep, logo); // ✅ SOLO para este reporte
      }

      int C0 = 0; // A
      int C_LAST = 10; // I
      // Anchos (A..K)
      shRep.setColumnWidth(0, 12000); // Detalle
      shRep.setColumnWidth(1, 4200);  // Desde Fecha
      shRep.setColumnWidth(2, 3000);  // Desde Hora
      shRep.setColumnWidth(3, 4200);  // Hasta Fecha
      shRep.setColumnWidth(4, 3000);  // Hasta Hora
      shRep.setColumnWidth(5, 2400);  // Desc días
      shRep.setColumnWidth(6, 2400);  // Desc hor
      shRep.setColumnWidth(7, 2400);  // Desc min
      shRep.setColumnWidth(8, 2400);  // Saldo días
      shRep.setColumnWidth(9, 2400);  // Saldo hor
      shRep.setColumnWidth(10, 2400); // Saldo min
      for (int c = 11; c <= 20; c++) {
        shRep.setColumnWidth(c, 2400);
      }

      // Estilos específicos del reporte
      CellStyle stTitulo = estiloTitulo;
      CellStyle stFranja = ConfiguracionExcel.crearEstiloFranjaVerde(libro);
      CellStyle stCabGray = ConfiguracionExcel.crearEstiloCabeceraGris(libro);
      CellStyle stPeriodo = ConfiguracionExcel.crearEstiloPeriodo(libro);
      CellStyle stHeadBlue = ConfiguracionExcel.crearEstiloHeaderAzul(libro);
      CellStyle stSubBlue = ConfiguracionExcel.crearEstiloSubHeaderAzul(libro);
      CellStyle stCellC = estiloCentroBorde;
      CellStyle stCellL = estiloIzqBorde;
      CellStyle stBoxTitle = ConfiguracionExcel.crearEstiloCajaTitulo(libro);
      CellStyle stBoxCell = ConfiguracionExcel.crearEstiloCajaDato(libro);
      CellStyle stPendTit = ConfiguracionExcel.crearEstiloPendientesTitulo(libro);
      CellStyle stTotal = ConfiguracionExcel.crearEstiloTotal(libro);

      int row = 0;

      // Títulos
      mergeSafeNoBorder(shRep, row, row, C0, C_LAST, stTitulo);
      UtilExcel.establecerTexto(shRep, row++, C0, UtilExcel.aMayusculasSeguras(safe(request.getEmpresa())), stTitulo);

      mergeSafeNoBorder(shRep, row, row, C0, C_LAST, stTitulo);
      UtilExcel.establecerTexto(shRep, row++, C0, "KARDEX VACACIONES", stTitulo);

      mergeSafeNoBorder(shRep, row, row, C0, C_LAST, stTitulo);
      UtilExcel.establecerTexto(shRep, row++, C0, "FECHA DE CORTE: " + safe(request.getFechaCorte()), stTitulo);
      row++; // espacio

      for (Map.Entry<String, List<KardexEmpleadoDTO>> entry : grupos.entrySet()) {
        List<KardexEmpleadoDTO> empsGrupo = entry.getValue();
        if (empsGrupo == null || empsGrupo.isEmpty())
          continue;

        for (KardexEmpleadoDTO emp : empsGrupo) {
          if (emp == null || emp.getPeriodos() == null || emp.getPeriodos().isEmpty())
            continue;

          String empNom = (safe(emp.getApellido()) + " " + safe(emp.getNombre())).trim();

          // FRANJA (según criterio, excepto 'e')
          String franja = textoFranjaEmpleado(criterio, emp);
          if (!franja.isBlank()) {
            mergeSafe(shRep, row, row, C0, C_LAST, stFranja);
            UtilExcel.establecerTexto(shRep, row++, C0, franja, stFranja);
          }

          int rCabIni = row;

          // fila 1 gris (CIUDAD / CC / COD)
          mergeSafeNoBorder(shRep, row, row, 0, 3, stCabGray);
          mergeSafeNoBorder(shRep, row, row, 4, 7, stCabGray);
          mergeSafeNoBorder(shRep, row, row, 8, 10, stCabGray);

          UtilExcel.establecerTexto(shRep, row, 0, "CIUDAD: " + safe(emp.getCiudad()), stCabGray);
          UtilExcel.establecerTexto(shRep, row, 4, "C.C.: " + safe(emp.getIdentificacion()), stCabGray);
          UtilExcel.establecerTexto(shRep, row, 8, "COD: " + safe(emp.getCodigo()), stCabGray);
          row++;

          // fila 2 gris (EMPLEADO)
          mergeSafeNoBorder(shRep, row, row, C0, C_LAST, stCabGray);
          UtilExcel.establecerTexto(shRep, row, C0, "EMPLEADO: " + empNom, stCabGray);
          row++;

          int rCabFin = row - 1;

          // ✅ 1) quitar bordes internos (2 filas x 0..8)
          limpiarBordesEnRegion(shRep, rCabIni, rCabFin, 0, 10);
          CellRangeAddress regionCab = new CellRangeAddress(rCabIni, rCabFin, 0, 10);
          bordeExternoGrueso(shRep, regionCab);

          for (KardexPeriodoDTO per : emp.getPeriodos()) {
            if (per == null)
              continue;

            String est = "ACTIVO".equalsIgnoreCase(safe(per.getEstado_periodo())) ? "Activo" : "Inactivo";

            // Periodo (NO merge 6..6)
            mergeSafe(shRep, row, row, 0, 1, stPeriodo);
            mergeSafe(shRep, row, row, 2, 3, stPeriodo);
            mergeSafe(shRep, row, row, 4, 5, stPeriodo);
            mergeSafe(shRep, row, row, 6, 7, stPeriodo);
            mergeSafe(shRep, row, row, 8, 10, stPeriodo);

            UtilExcel.establecerTexto(shRep, row, 0, "F.INICIO " + safe(per.getF_ingreso()), stPeriodo);
            UtilExcel.establecerTexto(shRep, row, 2, "F.FIN " + safe(per.getF_salida()), stPeriodo);
            UtilExcel.establecerTexto(shRep, row, 4, "F.CARGA " + safe(per.getF_carga()), stPeriodo);
            UtilExcel.establecerTexto(shRep, row, 6, "DÍAS ACUM " + fmtDec2(per.getDias_acum_decimal()), stPeriodo);
            UtilExcel.establecerTexto(shRep, row, 8, "ESTADO " + est, stPeriodo);
            row++;

            int rowInicioPeriodo = row;

            // Header 1 mov
            mergeSafe(shRep, row, row + 1, 0, 0, stHeadBlue);
            mergeSafe(shRep, row, row, 1, 2, stHeadBlue);
            mergeSafe(shRep, row, row, 3, 4, stHeadBlue);
            mergeSafe(shRep, row, row, 5, 7, stHeadBlue);
            mergeSafe(shRep, row, row, 8, 10, stHeadBlue);

            UtilExcel.establecerTexto(shRep, row, 0, "Detalle", stHeadBlue);
            UtilExcel.establecerTexto(shRep, row, 1, "Desde", stHeadBlue);
            UtilExcel.establecerTexto(shRep, row, 3, "Hasta", stHeadBlue);
            UtilExcel.establecerTexto(shRep, row, 5, "Descuento", stHeadBlue);
            UtilExcel.establecerTexto(shRep, row, 8, "Saldo", stHeadBlue);
            row++;

            // Header 2 mov
            Row h2 = UtilExcel.asegurarFila(shRep, row);
            UtilExcel.establecerTexto(h2, 1, "Fecha", stSubBlue);
            UtilExcel.establecerTexto(h2, 2, "Hora", stSubBlue);
            UtilExcel.establecerTexto(h2, 3, "Fecha", stSubBlue);
            UtilExcel.establecerTexto(h2, 4, "Hora", stSubBlue);
            UtilExcel.establecerTexto(h2, 5, "Días", stSubBlue);
            UtilExcel.establecerTexto(h2, 6, "Hor", stSubBlue);
            UtilExcel.establecerTexto(h2, 7, "Min", stSubBlue);
            UtilExcel.establecerTexto(h2, 8, "Días", stSubBlue);
            UtilExcel.establecerTexto(h2, 9, "Hor", stSubBlue);
            UtilExcel.establecerTexto(h2, 10, "Min", stSubBlue);
            row++;

            List<KardexMovimientoDTO> movs = per.getMovimientos();
            if (movs == null || movs.isEmpty()) {
              mergeSafe(shRep, row, row, 0, 10, stCellC);
              UtilExcel.establecerTexto(shRep, row++, 0, "Sin movimientos autorizados hasta la fecha de corte",
                  stCellC);
            } else {
              for (KardexMovimientoDTO m : movs) {
                Row rr = UtilExcel.asegurarFila(shRep, row++);
                DHMDTO d = m.getDescuento_dhm();
                DHMDTO s = m.getSaldo_dhm();

                UtilExcel.establecerTexto(rr, 0, safe(m.getDetalle()), stCellL);
                UtilExcel.establecerTexto(rr, 1, safe(m.getFecha_inicio()), stCellC);
                UtilExcel.establecerTexto(rr, 2, safe(m.getHora_inicio()), stCellC);
                UtilExcel.establecerTexto(rr, 3, safe(m.getFecha_final()), stCellC);
                UtilExcel.establecerTexto(rr, 4, safe(m.getHora_fin()), stCellC);

                UtilExcel.establecerValor(rr, 5, int0(d != null ? d.getDias() : 0), stCellC);
                UtilExcel.establecerValor(rr, 6, int0(d != null ? d.getHoras() : 0), stCellC);
                UtilExcel.establecerValor(rr, 7, int0(d != null ? d.getMinutos() : 0), stCellC);

                UtilExcel.establecerValor(rr, 8, int0(s != null ? s.getDias() : 0), stCellC);
                UtilExcel.establecerValor(rr, 9, int0(s != null ? s.getHoras() : 0), stCellC);
                UtilExcel.establecerValor(rr, 10, int0(s != null ? s.getMinutos() : 0), stCellC);
              }
            }

            // Bloque antigüedad / proporcional / liquidación
            Row bx1 = UtilExcel.asegurarFila(shRep, row);
            UtilExcel.establecerTexto(bx1, 0, "Días Antigüedad", stBoxTitle);

            mergeSafe(shRep, row, row, 5, 7, stBoxTitle);
            UtilExcel.establecerTexto(bx1, 5, "Proporcional", stBoxTitle);

            mergeSafe(shRep, row, row, 8, 10, stBoxTitle);
            UtilExcel.establecerTexto(bx1, 8, "Liquidación", stBoxTitle);
            row++;

            Row bx2 = UtilExcel.asegurarFila(shRep, row);
            UtilExcel.establecerTexto(bx2, 0, String.valueOf(int0(per.getDias_antiguedad())), stBoxCell);

            UtilExcel.establecerTexto(bx2, 5, "Días", stSubBlue);
            UtilExcel.establecerTexto(bx2, 6, "Hor", stSubBlue);
            UtilExcel.establecerTexto(bx2, 7, "Min", stSubBlue);

            UtilExcel.establecerTexto(bx2, 8, "Días", stSubBlue);
            UtilExcel.establecerTexto(bx2, 9, "Hor", stSubBlue);
            UtilExcel.establecerTexto(bx2, 10, "Min", stSubBlue);
            row++;

            DHMDTO pd = per.getProporcional_dhm();
            DHMDTO ld = per.getLiquidacion_dhm();

            Row bx3 = UtilExcel.asegurarFila(shRep, row++);
            UtilExcel.establecerTexto(bx3, 0, "", stBoxCell);

            UtilExcel.establecerValor(bx3, 5, int0(pd != null ? pd.getDias() : 0), stBoxCell);
            UtilExcel.establecerValor(bx3, 6, int0(pd != null ? pd.getHoras() : 0), stBoxCell);
            UtilExcel.establecerValor(bx3, 7, int0(pd != null ? pd.getMinutos() : 0), stBoxCell);

            UtilExcel.establecerValor(bx3, 8, int0(ld != null ? ld.getDias() : 0), stBoxCell);
            UtilExcel.establecerValor(bx3, 9, int0(ld != null ? ld.getHoras() : 0), stBoxCell);
            UtilExcel.establecerValor(bx3, 10, int0(ld != null ? ld.getMinutos() : 0), stBoxCell);

            row++;

            escribirDistribucionExcelLateral(
                shRep,
                rowInicioPeriodo,
                per,
                stPendTit,
                stHeadBlue,
                stSubBlue,
                stCellC
            );
            
            row++;

            // Pendientes (si hay)
            List<KardexPendienteDTO> pend = per.getPendientes();
            if (pend != null && !pend.isEmpty()) {

              mergeSafe(shRep, row, row, 0, 8, stPendTit);
              UtilExcel.establecerTexto(shRep, row++, 0,
                  "Solicitudes de permisos y vacaciones (Pendientes, Pre-Autorizadas)", stPendTit);

                  mergeSafe(shRep, row, row + 1, 0, 0, stHeadBlue);
                  mergeSafe(shRep, row, row, 1, 2, stHeadBlue);
                  mergeSafe(shRep, row, row, 3, 4, stHeadBlue);
                  mergeSafe(shRep, row, row, 5, 7, stHeadBlue);
                  mergeSafe(shRep, row, row + 1, 8, 10, stHeadBlue);

                  UtilExcel.establecerTexto(shRep, row, 0, "Detalle", stHeadBlue);
                  UtilExcel.establecerTexto(shRep, row, 1, "Desde", stHeadBlue);
                  UtilExcel.establecerTexto(shRep, row, 3, "Hasta", stHeadBlue);
                  UtilExcel.establecerTexto(shRep, row, 5, "Tiempo", stHeadBlue);
                  UtilExcel.establecerTexto(shRep, row, 8, "Estado", stHeadBlue);
                  row++;

                  Row hp2 = UtilExcel.asegurarFila(shRep, row++);
                  UtilExcel.establecerTexto(hp2, 1, "Fecha", stSubBlue);
                  UtilExcel.establecerTexto(hp2, 2, "Hora", stSubBlue);
                  UtilExcel.establecerTexto(hp2, 3, "Fecha", stSubBlue);
                  UtilExcel.establecerTexto(hp2, 4, "Hora", stSubBlue);
                  UtilExcel.establecerTexto(hp2, 5, "Días", stSubBlue);
                  UtilExcel.establecerTexto(hp2, 6, "Hor", stSubBlue);
                  UtilExcel.establecerTexto(hp2, 7, "Min", stSubBlue);

              double totalDec = 0.0;

              for (KardexPendienteDTO x : pend) {
                Row pr = UtilExcel.asegurarFila(shRep, row++);
                DHMDTO t = pendienteDHM(emp, x);
                totalDec += toDecimalDias(x.getDias(), x.getMinutos_totales(), int0(emp.getMin_por_dia()));

                UtilExcel.establecerTexto(pr, 0, safe(x.getDetalle()), stCellL);
                UtilExcel.establecerTexto(pr, 1, safe(x.getFecha_inicio()), stCellC);
                UtilExcel.establecerTexto(pr, 2, safe(x.getHora_inicio()), stCellC);
                UtilExcel.establecerTexto(pr, 3, safe(x.getFecha_final()), stCellC);
                UtilExcel.establecerTexto(pr, 4, safe(x.getHora_fin()), stCellC);

                UtilExcel.establecerValor(pr, 5, int0(t.getDias()), stCellC);
                UtilExcel.establecerValor(pr, 6, int0(t.getHoras()), stCellC);
                UtilExcel.establecerValor(pr, 7, int0(t.getMinutos()), stCellC);

                mergeSafe(shRep, row - 1, row - 1, 8, 10, stCellC);
                UtilExcel.establecerTexto(shRep, row - 1, 8, safe(x.getEstado_texto()), stCellC);
              }

              DHMDTO total = decimalToDHM(totalDec, int0(emp.getMin_por_dia()));
              Row tr = UtilExcel.asegurarFila(shRep, row++);

              mergeSafe(shRep, row - 1, row - 1, 0, 4, stTotal);
              UtilExcel.establecerTexto(shRep, row - 1, 0, "Total estimado:", stTotal);

              UtilExcel.establecerValor(tr, 5, int0(total.getDias()), stTotal);
              UtilExcel.establecerValor(tr, 6, int0(total.getHoras()), stTotal);
              UtilExcel.establecerValor(tr, 7, int0(total.getMinutos()), stTotal);

              mergeSafe(shRep, row - 1, row - 1, 8, 10, stTotal);
              UtilExcel.establecerTexto(shRep, row - 1, 8, "", stTotal);

              row++;
            }

            row += 2;
          }

          row += 2;
        }
      }

      shRep.createFreezePane(0, 4);

      libro.write(baos);
      return baos.toByteArray();

    } catch (IllegalArgumentException e) {
      throw e;
    } catch (Exception e) {
      throw new ReportBuildException("No se pudo generar KardexVacaciones.xlsx", e);
    }
  }

  // =========================================================================================
  // Helpers PDF (celdas)
  // =========================================================================================
  private PdfPCell hCell(String text, Color bg) {
    return ReporteUtil.crearCelda(text, ReporteUtil.fuenteEncabezado(), bg);
  }

  private PdfPCell hCell(String text, Color bg, int colspan, int rowspan) {
    PdfPCell c = ReporteUtil.crearCelda(text, ReporteUtil.fuenteEncabezado(), bg);
    c.setColspan(colspan);
    c.setRowspan(rowspan);
    c.setHorizontalAlignment(Element.ALIGN_CENTER);
    c.setVerticalAlignment(Element.ALIGN_MIDDLE);
    return c;
  }

  private PdfPCell cellCenter(String text, Color bg) {
    PdfPCell c = new PdfPCell(new Phrase(safe(text), ReporteUtil.fuenteTablaData())); // <-- NORMAL
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
    PdfPCell c = new PdfPCell(new Phrase(safe(text), ReporteUtil.fuenteTablaData())); // <-- NORMAL
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
    // etiqueta en negrita (encabezado de tabla data)
    contenido.add(new Chunk(safe(etiqueta) + " ", ReporteUtil.fuenteEncabezadoTablaData()));
    // valor normal (data)
    contenido.add(new Chunk(safe(valor), ReporteUtil.fuenteTablaData()));

    PdfPCell celda = new PdfPCell(contenido);
    celda.setBackgroundColor(fondo);
    celda.setPadding(2f);
    celda.setBorder(Rectangle.NO_BORDER);
    return celda;
  }

  // =========================================================================================
  // Helpers negocio (pendientes DHM igual que frontend)
  // =========================================================================================
  private DHMDTO pendienteDHM(KardexEmpleadoDTO emp, KardexPendienteDTO x) {
    int minPorDia = int0(emp != null ? emp.getMin_por_dia() : 0);
    double dec = toDecimalDias(x != null ? x.getDias() : "0", x != null ? x.getMinutos_totales() : 0, minPorDia);
    return decimalToDHM(dec, minPorDia);
  }

  private double toDecimalDias(String dias, Integer minutos, int minPorDia) {
    double d = 0.0;
    try {
      d = Double.parseDouble(safe(dias));
    } catch (Exception ignore) {
      d = 0.0;
    }
    int m = (minutos == null) ? 0 : minutos;
    if (minPorDia <= 0)
      return d;
    return d + (m / (double) minPorDia);
  }

  private DHMDTO decimalToDHM(double decimalDias, int minPorDia) {
    DHMDTO out = new DHMDTO();

    if (minPorDia <= 0) {
      int dias = (decimalDias < 0) ? -((int) Math.floor(Math.abs(decimalDias))) : (int) Math.floor(decimalDias);
      out.setDias(dias);
      out.setHoras(0);
      out.setMinutos(0);
      return out;
    }

    int sign = decimalDias < 0 ? -1 : 1;
    double abs = Math.abs(decimalDias);

    int diasAbs = (int) Math.floor(abs);
    double resto = abs - diasAbs;

    int totalMinAbs = (int) Math.round(resto * minPorDia);
    int horasAbs = totalMinAbs / 60;
    int minAbs = totalMinAbs % 60;

    out.setDias(sign * diasAbs);
    out.setHoras(sign * horasAbs);
    out.setMinutos(sign * minAbs);
    return out;
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

  private int int0(Integer v) {
    return (v == null) ? 0 : v;
  }

  private String fmtDec2(Double n) {
    if (n == null)
      return "0,00";
    return String.format(java.util.Locale.US, "%.2f", n).replace('.', ',');
  }

  // =========================================================================================
  // ✅ Agrupación y franja por criterio
  // =========================================================================================
  private String grupoNombre(String criterio, KardexEmpleadoDTO e) {
    if (e == null)
      return "";
    String c = safe(criterio).toLowerCase();
    switch (c) {
      case "s":
        return pickEmpSucursal(e);
      case "r":
        return pickEmpRegimen(e);
      case "d":
        return pickEmpDepartamento(e);
      case "c":
        return pickEmpCargo(e);
      case "e":
        return (safe(e.getApellido()) + " " + safe(e.getNombre())).trim();
      default:
        return "";
    }
  }

  private String textoFranjaEmpleado(String criterio, KardexEmpleadoDTO e) {
    String c = safe(criterio).toLowerCase();
    if ("e".equals(c))
      return ""; // en "empleados" no necesitas franja

    String val = safe(grupoNombre(c, e));
    if (val.isBlank())
      return "";

    switch (c) {
      case "s":
        return "SUCURSAL: " + val;
      case "r":
        return "RÉGIMEN: " + val;
      case "d":
        return "DEPARTAMENTO: " + val;
      case "c":
        return "CARGO: " + val;
      default:
        return "";
    }
  }

  // ✅ helper para poner al final del service (NO toca UtilExcel)
  // ✅ helper mejorado (NO toca UtilExcel)
  private void mergeSafe(XSSFSheet sh, int r1, int r2, int c1, int c2, CellStyle estilo) {
    // 1 sola celda => solo asegurar celda/estilo y salir
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

    // Evitar duplicados si por alguna razón se llama 2 veces al mismo merge
    // (opcional, pero ayuda a no romper el XLSX)
    for (int i = 0; i < sh.getNumMergedRegions(); i++) {
      CellRangeAddress ex = sh.getMergedRegion(i);
      if (ex.formatAsString().equals(region.formatAsString())) {
        // ya existe
        UtilExcel.aplicarEstiloARegion(sh, rr1, rr2, cc1, cc2, estilo, true);
        aplicarBordeRegion(sh, region);
        return;
      }
    }

    sh.addMergedRegion(region);

    // 1) aplicar estilo a todo el rango (crea celdas faltantes)
    UtilExcel.aplicarEstiloARegion(sh, rr1, rr2, cc1, cc2, estilo, true);

    // 2) dibujar bordes externos del rango mergeado (cierra “cajas” y headers)
    aplicarBordeRegion(sh, region);
  }

  private void aplicarBordeRegion(Sheet sh, CellRangeAddress region) {
    RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
    RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
    RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
    RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);

    // color borde (ajústalo si quieres gris)
    short col = IndexedColors.GREY_50_PERCENT.getIndex();
    RegionUtil.setTopBorderColor(col, region, sh);
    RegionUtil.setBottomBorderColor(col, region, sh);
    RegionUtil.setLeftBorderColor(col, region, sh);
    RegionUtil.setRightBorderColor(col, region, sh);
  }

  // =========================================================================================
  // ✅ Tomar nombre real del empleado (prioriza campo “bonito” y cae al name_*)
  // =========================================================================================
  private String pickEmpSucursal(KardexEmpleadoDTO e) {
    String a = safe(e.getSucursal());
    if (!a.isBlank())
      return a;
    return safe(e.getName_suc());
  }

  private String pickEmpDepartamento(KardexEmpleadoDTO e) {
    String a = safe(e.getDepartamento());
    if (!a.isBlank())
      return a;
    return safe(e.getName_dep());
  }

  private String pickEmpCargo(KardexEmpleadoDTO e) {
    String a = safe(e.getCargo());
    if (!a.isBlank())
      return a;
    return safe(e.getName_cargo());
  }

  private String pickEmpRegimen(KardexEmpleadoDTO e) {
    String a = safe(e.getRegimen());
    if (!a.isBlank())
      return a;
    return safe(e.getName_regimen());
  }

  private PdfPCell wrapNoSpace(PdfPTable t) {
    PdfPCell c = new PdfPCell(t);
    c.setPadding(0f);
    c.setBorder(Rectangle.NO_BORDER);
    return c;
  }

  private void noSpace(PdfPTable t) {
    t.setSpacingBefore(0f);
    t.setSpacingAfter(0f);
  }

  private void insertarLogoKardexA1A4(Workbook wb, Sheet hoja, byte[] imagenBytes) {
    if (imagenBytes == null || imagenBytes.length == 0)
      return;

    int idx = wb.addPicture(imagenBytes, Workbook.PICTURE_TYPE_PNG);

    Drawing<?> drawing = hoja.createDrawingPatriarch();
    CreationHelper helper = wb.getCreationHelper();
    ClientAnchor anchor = helper.createClientAnchor();

    // A1:A4
    anchor.setCol1(0); // A
    anchor.setCol2(1); // hasta antes de B
    anchor.setRow1(0); // fila 1
    anchor.setRow2(4); // fila 4

    // ✅ ancho EXACTO solo A
    anchor.setDx1(0);
    anchor.setDx2(0);

    // alto completo del rango
    anchor.setDy1(0);
    anchor.setDy2(0);

    anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

    drawing.createPicture(anchor, idx);
  }

  // ✅ merge SIN borde (para títulos superiores)
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

    // aplica el estilo a todo el rango para que el fondo/fuente sea uniforme,
    // pero NO dibuja borde con RegionUtil
    UtilExcel.aplicarEstiloARegion(sh, rr1, rr2, cc1, cc2, estilo, true);

    // ✅ quitar bordes del estilo SOLO para estas celdas (por si el estilo trae
    // bordes)
    // OJO: esto NO modifica el estilo global, solo el borde visible en la región:
    RegionUtil.setBorderTop(BorderStyle.NONE, region, sh);
    RegionUtil.setBorderBottom(BorderStyle.NONE, region, sh);
    RegionUtil.setBorderLeft(BorderStyle.NONE, region, sh);
    RegionUtil.setBorderRight(BorderStyle.NONE, region, sh);
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

        // ⚠️ No modifiques el mismo CellStyle compartido (POI lo reutiliza)
        // => clonamos a un nuevo estilo y lo aplicamos
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

  private PdfPTable construirTituloDistribucionPDF() throws Exception {
    PdfPTable tituloTbl = new PdfPTable(1);
    tituloTbl.setWidthPercentage(100);
    noSpace(tituloTbl);

    PdfPCell titulo = new PdfPCell(
        new Phrase("Distribución de días", ReporteUtil.fuenteEncabezado()));
    titulo.setHorizontalAlignment(Element.ALIGN_LEFT);
    titulo.setVerticalAlignment(Element.ALIGN_MIDDLE);
    titulo.setBackgroundColor(new Color(238, 243, 255));
    titulo.setBorder(Rectangle.BOX);
    titulo.setBorderColor(new Color(207, 216, 255));
    titulo.setPadding(6f);

    tituloTbl.addCell(titulo);
    tituloTbl.setSpacingBefore(3f);
    tituloTbl.setSpacingAfter(4f);

    return tituloTbl;
  }

  private PdfPTable construirTablaDistribucionPDF(
      KardexPeriodoDTO per,
      Color azulHeader,
      Color azulSub) throws Exception {

    PdfPTable tabla = new PdfPTable(21);
    tabla.setWidthPercentage(100);

    float[] widths = new float[21];
    for (int i = 0; i < 7; i++) {
      widths[i * 3] = 1f;
      widths[i * 3 + 1] = 1f;
      widths[i * 3 + 2] = 1f;
    }
    tabla.setWidths(widths);
    noSpace(tabla);

    String[] dias = { "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };

    for (String dia : dias) {
      tabla.addCell(hCell(dia, azulHeader, 3, 1));
    }

    for (int i = 0; i < 7; i++) {
      tabla.addCell(hCell("Días", azulSub));
      tabla.addCell(hCell("Hor", azulSub));
      tabla.addCell(hCell("Min", azulSub));
    }

    SemanaDHMDTO total = sumarSemanaPeriodo(per);

    addDhmCellsPDF(tabla, total != null ? total.getLunes() : null);
    addDhmCellsPDF(tabla, total != null ? total.getMartes() : null);
    addDhmCellsPDF(tabla, total != null ? total.getMiercoles() : null);
    addDhmCellsPDF(tabla, total != null ? total.getJueves() : null);
    addDhmCellsPDF(tabla, total != null ? total.getViernes() : null);
    addDhmCellsPDF(tabla, total != null ? total.getSabado() : null);
    addDhmCellsPDF(tabla, total != null ? total.getDomingo() : null);

    tabla.setSpacingAfter(10f);
    return tabla;
  }
    
  private void addDhmCellsPDF(PdfPTable tabla, DHMDTO dhm) {
    tabla.addCell(cellCenter(String.valueOf(int0(dhm != null ? dhm.getDias() : 0)), Color.WHITE));
    tabla.addCell(cellCenter(String.valueOf(int0(dhm != null ? dhm.getHoras() : 0)), Color.WHITE));
    tabla.addCell(cellCenter(String.valueOf(int0(dhm != null ? dhm.getMinutos() : 0)), Color.WHITE));
  }

  private SemanaDHMDTO sumarSemanaPeriodo(KardexPeriodoDTO per) {
    SemanaDHMDTO out = semanaVacia();

    if (per == null || per.getMovimientos() == null) {
      return out;
    }

    for (KardexMovimientoDTO mov : per.getMovimientos()) {
      if (mov == null || mov.getSemana_descuento() == null) continue;

      SemanaDHMDTO s = mov.getSemana_descuento();

      acumularDhm(out.getLunes(), s.getLunes());
      acumularDhm(out.getMartes(), s.getMartes());
      acumularDhm(out.getMiercoles(), s.getMiercoles());
      acumularDhm(out.getJueves(), s.getJueves());
      acumularDhm(out.getViernes(), s.getViernes());
      acumularDhm(out.getSabado(), s.getSabado());
      acumularDhm(out.getDomingo(), s.getDomingo());
    }

    return out;
  }

  private SemanaDHMDTO semanaVacia() {
    SemanaDHMDTO s = new SemanaDHMDTO();
    s.setLunes(dhmVacio());
    s.setMartes(dhmVacio());
    s.setMiercoles(dhmVacio());
    s.setJueves(dhmVacio());
    s.setViernes(dhmVacio());
    s.setSabado(dhmVacio());
    s.setDomingo(dhmVacio());
    return s;
  }

  private DHMDTO dhmVacio() {
    DHMDTO d = new DHMDTO();
    d.setDias(0);
    d.setHoras(0);
    d.setMinutos(0);
    return d;
  }

  private void acumularDhm(DHMDTO destino, DHMDTO origen) {
    if (destino == null || origen == null) return;

    destino.setDias(int0(destino.getDias()) + int0(origen.getDias()));
    destino.setHoras(int0(destino.getHoras()) + int0(origen.getHoras()));
    destino.setMinutos(int0(destino.getMinutos()) + int0(origen.getMinutos()));
  }

  private void escribirDistribucionExcelLateral(
      XSSFSheet sh,
      int rowInicioPeriodo,
      KardexPeriodoDTO per,
      CellStyle stTitulo,
      CellStyle stHeadBlue,
      CellStyle stSubBlue,
      CellStyle stCellC) {

    final int DIST_START = 12; // M

    SemanaDHMDTO total = sumarSemanaPeriodo(per);

    // fila 1: título (misma altura que fila período)
    mergeSafe(sh, rowInicioPeriodo - 1, rowInicioPeriodo - 1, DIST_START, DIST_START + 20, stTitulo);
    UtilExcel.establecerTexto(sh, rowInicioPeriodo - 1, DIST_START, "Distribución de días", stTitulo);

    // fila 2: días de la semana (misma altura que Desde/Hasta/Descuento/Saldo)
    String[] dias = { "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo" };
    int col = DIST_START;
    for (String dia : dias) {
      mergeSafe(sh, rowInicioPeriodo, rowInicioPeriodo, col, col + 2, stHeadBlue);
      UtilExcel.establecerTexto(sh, rowInicioPeriodo, col, dia, stHeadBlue);
      col += 3;
    }

    // fila 3: subcabeceras (misma altura que Fecha/Hora/Días/Hor/Min)
    Row sub = UtilExcel.asegurarFila(sh, rowInicioPeriodo + 1);
    col = DIST_START;
    for (int i = 0; i < 7; i++) {
      UtilExcel.establecerTexto(sub, col++, "Días", stSubBlue);
      UtilExcel.establecerTexto(sub, col++, "Hor", stSubBlue);
      UtilExcel.establecerTexto(sub, col++, "Min", stSubBlue);
    }

    // fila 4: valores
    Row data = UtilExcel.asegurarFila(sh, rowInicioPeriodo + 2);
    col = DIST_START;
    col = escribirDhmExcel(data, col, total != null ? total.getLunes() : null, stCellC);
    col = escribirDhmExcel(data, col, total != null ? total.getMartes() : null, stCellC);
    col = escribirDhmExcel(data, col, total != null ? total.getMiercoles() : null, stCellC);
    col = escribirDhmExcel(data, col, total != null ? total.getJueves() : null, stCellC);
    col = escribirDhmExcel(data, col, total != null ? total.getViernes() : null, stCellC);
    col = escribirDhmExcel(data, col, total != null ? total.getSabado() : null, stCellC);
    col = escribirDhmExcel(data, col, total != null ? total.getDomingo() : null, stCellC);
  }
    
  private int escribirDhmExcel(Row row, int col, DHMDTO dhm, CellStyle style) {
    UtilExcel.establecerValor(row, col++, int0(dhm != null ? dhm.getDias() : 0), style);
    UtilExcel.establecerValor(row, col++, int0(dhm != null ? dhm.getHoras() : 0), style);
    UtilExcel.establecerValor(row, col++, int0(dhm != null ? dhm.getMinutos() : 0), style);
    return col;
  }

}