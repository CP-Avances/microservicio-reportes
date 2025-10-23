package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.RegimenLaboral.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.casapazmino.microservicio_reportes.util.ReportBuildException;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteRegimenesService {

        // METODO QUE GENERA EL PDF
        public byte[] generarReporteRegimenesPDF(ReporteRegimenesRequest request) {

        // ➊ DRY: constantes locales (no cambia look & feel)
        final String TITULO = "RÉGIMEN LABORAL";
        final float[] WIDTHS_CABECERA = { 2.5f, 3f, 2.5f, 2f };
        final float[] WIDTHS_CONTENEDOR = { 3f, 3f, 3f };
        final float[] WIDTHS_DOBLE = { 2.5f, 1.5f };
        final int WIDTH_PERCENT_100 = 100;
        final float ESPACIO_ANTES_BLOQUE = 10f;
        final float ESPACIO_DESPUES_BLOQUE = 10f;
        final float ESPACIO_DESPUES_CABECERA = 6f;

        final Color COLOR_PRIMARIO   = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
        final Color COLOR_SECUNDARIO = ReporteUtil.convertirHexAColor(request.getColorSecundario());
        final Color COLOR_ZEBRA      = ReporteUtil.colorZebraClaro();

        Document document = null;
        PdfWriter writer = null;
        ByteArrayOutputStream baos = null;

        try {
                // 1) Inicialización
                baos = new ByteArrayOutputStream();
                document = new Document(PageSize.A4.rotate());
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

                // Iteración de regímenes
                if (request.getRegimenes() != null) {
                for (RegimenDTO reg : request.getRegimenes()) {
                        int contador = 0;

                        // Tabla cabecera del régimen
                        PdfPTable tablaCabecera = new PdfPTable(4);
                        tablaCabecera.setWidthPercentage(WIDTH_PERCENT_100);
                        tablaCabecera.setWidths(WIDTHS_CABECERA);

                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("PAÍS: ", reg.getPais(), COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN: ", reg.getDescripcion(), COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("CONTINUIDAD LABORAL: ",
                                (reg.getContinuidad_laboral() ? "SÍ" : "NO"), COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("CÓDIGO: ",
                                String.valueOf(reg.getId()), COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("PERIODO LABORAL: ",
                                reg.getMes_periodo() + " Meses", COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("DÍAS POR MES: ",
                                String.valueOf(reg.getDias_mes()), COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("TIEMPO MÍNIMO: ",
                                (reg.getTrabajo_minimo_mes() > 0
                                        ? reg.getTrabajo_minimo_mes() + " Meses"
                                        : reg.getTrabajo_minimo_horas() + " Horas"),
                                COLOR_PRIMARIO));
                        tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("ANTIGÜEDAD LABORAL: ",
                                (reg.getAntiguedad() ? "SÍ" : "NO"), COLOR_PRIMARIO));

                        // Contenedor para borde externo
                        PdfPTable tablaContenedora = new PdfPTable(1);
                        tablaContenedora.setWidthPercentage(WIDTH_PERCENT_100);
                        PdfPCell contenedor1 = new PdfPCell(tablaCabecera);
                        contenedor1.setBorder(Rectangle.BOX);
                        tablaContenedora.addCell(contenedor1);
                        tablaContenedora.setSpacingBefore(ESPACIO_ANTES_BLOQUE);
                        tablaContenedora.setSpacingAfter(ESPACIO_DESPUES_CABECERA);
                        document.add(tablaContenedora);

                        // Tabla 1: Configuración de Vacaciones
                        PdfPTable configVac = new PdfPTable(2);
                        configVac.setWidthPercentage(WIDTH_PERCENT_100);
                        configVac.setWidths(WIDTHS_DOBLE);
                        configVac.addCell(ReporteUtil.crearCelda(
                                "CONFIGURACIÓN DE VACACIONES",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO, 1, 2));

                        configVac.addCell(ReporteUtil.celdaCentro("DÍAS HÁBILES", COLOR_SECUNDARIO));
                        configVac.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getVacacion_dias_laboral()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        configVac.addCell(ReporteUtil.celdaCentro("DÍAS LIBRES", COLOR_SECUNDARIO));
                        configVac.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getVacacion_dias_libre()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        configVac.addCell(ReporteUtil.celdaCentro("DÍAS CALENDARIO", COLOR_SECUNDARIO));
                        configVac.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getVacacion_dias_calendario()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        configVac.addCell(ReporteUtil.celdaCentro("ACUMULA VACACIONES", COLOR_SECUNDARIO));
                        configVac.addCell(ReporteUtil.celdaCentro(
                                reg.getAcumular() ? "SÍ" : "NO",
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        if (reg.getAcumular()) {
                        configVac.addCell(ReporteUtil.celdaCentro("MÁXIMO DÍAS ACUMULABLES", COLOR_SECUNDARIO));
                        configVac.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getDias_maximo_acumulacion()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));
                        }

                        configVac.addCell(ReporteUtil.celdaCentro("VACACIONES POR PERÍODOS", COLOR_SECUNDARIO));
                        configVac.addCell(ReporteUtil.celdaCentro(
                                reg.getVacacion_divisible() ? "SÍ" : "NO",
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        if (reg.getVacacion_divisible()
                                && reg.getPeriodos_vacacionales() != null
                                && !reg.getPeriodos_vacacionales().isEmpty()) {
                        for (PeriodoVacacionalDTO p : reg.getPeriodos_vacacionales()) {
                                configVac.addCell(ReporteUtil.celdaCentro(
                                        p.getDescripcion(),
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));
                                configVac.addCell(ReporteUtil.celdaCentro(
                                        p.getDias_vacacion() + " días",
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));
                        }
                        }

                        // Tabla 2: Vacaciones ganadas
                        PdfPTable vacGanadas = new PdfPTable(2);
                        vacGanadas.setWidthPercentage(WIDTH_PERCENT_100);
                        vacGanadas.setWidths(WIDTHS_DOBLE);
                        vacGanadas.addCell(ReporteUtil.crearCelda(
                                "VACACIONES GANADAS",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO, 1, 2));

                        vacGanadas.addCell(ReporteUtil.celdaCentro("POR MES (HÁBILES)", COLOR_SECUNDARIO));
                        vacGanadas.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getVacacion_dias_laboral_mes()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        vacGanadas.addCell(ReporteUtil.celdaCentro("POR MES (CALENDARIO)", COLOR_SECUNDARIO));
                        vacGanadas.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getVacacion_dias_calendario_mes()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        vacGanadas.addCell(ReporteUtil.celdaCentro("POR DÍA (HÁBILES)", COLOR_SECUNDARIO));
                        vacGanadas.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getLaboral_dias()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        vacGanadas.addCell(ReporteUtil.celdaCentro("POR DÍA (CALENDARIO)", COLOR_SECUNDARIO));
                        vacGanadas.addCell(ReporteUtil.celdaCentro(
                                String.valueOf(reg.getCalendario_dias()),
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        // Tabla 3: Configuración de Antigüedad
                        PdfPTable antiguedad = new PdfPTable(2);
                        antiguedad.setWidthPercentage(WIDTH_PERCENT_100);
                        antiguedad.setWidths(WIDTHS_DOBLE);
                        antiguedad.addCell(ReporteUtil.crearCelda(
                                "CONFIGURACIÓN DE ANTIGÜEDAD",
                                ReporteUtil.fuenteEncabezadoTablaData(),
                                COLOR_PRIMARIO, 1, 2));

                        if (Boolean.TRUE.equals(reg.getAntiguedad())) {
                        if (Boolean.TRUE.equals(reg.getAntiguedad_fija())) {
                                antiguedad.addCell(ReporteUtil.celdaCentro("TIPO", COLOR_SECUNDARIO));
                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                        "FIJA",
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                                antiguedad.addCell(ReporteUtil.celdaCentro("AÑOS ANTIGÜEDAD", COLOR_SECUNDARIO));
                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                        String.valueOf(reg.getAnio_antiguedad()),
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                                antiguedad.addCell(ReporteUtil.celdaCentro("DÍAS ADICIONALES", COLOR_SECUNDARIO));
                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                        String.valueOf(reg.getDias_antiguedad()),
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                        } else if (Boolean.TRUE.equals(reg.getAntiguedad_variable())
                                && reg.getRangos_antiguedad() != null
                                && !reg.getRangos_antiguedad().isEmpty()) {

                                antiguedad.addCell(ReporteUtil.celdaCentro("TIPO", COLOR_SECUNDARIO));
                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                        "VARIABLE",
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));

                                for (RangoAntiguedadDTO r : reg.getRangos_antiguedad()) {
                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                        "Desde " + r.getAnio_desde() + " hasta " + r.getAnio_hasta() + " años",
                                        COLOR_SECUNDARIO));
                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                        r.getDias_antiguedad() + " días",
                                        (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE));
                                }
                        }
                        } else {
                        PdfPCell celdaNoAplica = ReporteUtil.celdaCentro(
                                "NO APLICA",
                                (contador++ % 2 == 0) ? COLOR_ZEBRA : Color.WHITE);
                        celdaNoAplica.setColspan(2);
                        antiguedad.addCell(celdaNoAplica);
                        }

                        // Normalizamos filas para alinear alturas entre las 3 tablas
                        int filasVac = configVac.size();
                        int filasGanadas = vacGanadas.size();
                        int filasAntig = antiguedad.size();
                        int maxFilas = Math.max(filasVac, Math.max(filasGanadas, filasAntig));

                        while (configVac.size() < maxFilas) {
                        PdfPCell c1 = ReporteUtil.celdaCentro("", Color.WHITE); c1.setBorder(Rectangle.NO_BORDER);
                        PdfPCell c2 = ReporteUtil.celdaCentro("", Color.WHITE); c2.setBorder(Rectangle.NO_BORDER);
                        configVac.addCell(c1); configVac.addCell(c2);
                        }
                        while (vacGanadas.size() < maxFilas) {
                        PdfPCell c1 = ReporteUtil.celdaCentro("", Color.WHITE); c1.setBorder(Rectangle.NO_BORDER);
                        PdfPCell c2 = ReporteUtil.celdaCentro("", Color.WHITE); c2.setBorder(Rectangle.NO_BORDER);
                        vacGanadas.addCell(c1); vacGanadas.addCell(c2);
                        }
                        while (antiguedad.size() < maxFilas) {
                        PdfPCell c1 = ReporteUtil.celdaCentro("", Color.WHITE); c1.setBorder(Rectangle.NO_BORDER);
                        PdfPCell c2 = ReporteUtil.celdaCentro("", Color.WHITE); c2.setBorder(Rectangle.NO_BORDER);
                        antiguedad.addCell(c1); antiguedad.addCell(c2);
                        }

                        // Contenedor 3-col
                        PdfPCell celdaConfigVac = new PdfPCell(configVac);
                        celdaConfigVac.setVerticalAlignment(Element.ALIGN_TOP);
                        celdaConfigVac.setPadding(0);
                        celdaConfigVac.setBorder(Rectangle.NO_BORDER);

                        PdfPCell celdaVacGanadas = new PdfPCell(vacGanadas);
                        celdaVacGanadas.setVerticalAlignment(Element.ALIGN_TOP);
                        celdaVacGanadas.setPadding(0);
                        celdaVacGanadas.setBorder(Rectangle.NO_BORDER);

                        PdfPCell celdaAntiguedad = new PdfPCell(antiguedad);
                        celdaAntiguedad.setVerticalAlignment(Element.ALIGN_TOP);
                        celdaAntiguedad.setPadding(0);
                        celdaAntiguedad.setBorder(Rectangle.NO_BORDER);

                        PdfPTable contenedor = new PdfPTable(3);
                        contenedor.setWidthPercentage(WIDTH_PERCENT_100);
                        contenedor.setWidths(WIDTHS_CONTENEDOR);
                        contenedor.addCell(celdaConfigVac);
                        contenedor.addCell(celdaVacGanadas);
                        contenedor.addCell(celdaAntiguedad);
                        contenedor.setSpacingAfter(ESPACIO_DESPUES_BLOQUE);

                        contador++;
                        document.add(contenedor);
                }
                }

                // 3) Cierre + retorno
                document.close();
                return baos.toByteArray();

        } catch (IllegalArgumentException e) {
                // Validaciones de helpers → el controller decidirá 400 si aplica
                throw e;
        } catch (Exception e) {
                // Fallo interno uniforme → 500
                throw new ReportBuildException("No se pudo generar ReporteRegimenes.pdf", e);
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
        // ======================= XLSX =======================
        public byte[] generarReporteRegimenesXLSX(ReporteRegimenesRequest request) {
                try (org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
                                ByteArrayOutputStream out = new ByteArrayOutputStream()) {

                        org.apache.poi.xssf.usermodel.XSSFSheet hoja = wb.createSheet("Régimen");

                        // 1) Logo estándar A1:B5 (igual que en otros módulos)
                        byte[] logo = com.casapazmino.microservicio_reportes.util.UtilExcel
                                        .decodificarImagenBase64(request.getLogoBase64());
                        com.casapazmino.microservicio_reportes.util.UtilExcel.insertarLogoEstandar(wb, hoja, logo);

                        // 2) Estilos
                        org.apache.poi.ss.usermodel.CellStyle estTitulo = com.casapazmino.microservicio_reportes.util.ConfiguracionExcel
                                        .crearEstiloTitulo(wb);
                        org.apache.poi.ss.usermodel.CellStyle estEnc = com.casapazmino.microservicio_reportes.util.ConfiguracionExcel
                                        .crearEstiloEncabezadoTabla(wb);
                        org.apache.poi.ss.usermodel.CellStyle estCentro = com.casapazmino.microservicio_reportes.util.ConfiguracionExcel
                                        .crearEstiloCentroConBorde(wb);
                        org.apache.poi.ss.usermodel.CellStyle estIzquierda = com.casapazmino.microservicio_reportes.util.ConfiguracionExcel
                                        .crearEstiloIzquierdaConBorde(wb);

                        // 3) Títulos (B1:Y1 y B2:Y2 → 25 columnas A..Y)
                        com.casapazmino.microservicio_reportes.util.UtilExcel.combinarCeldas(hoja, 0, 0, 1, 24); // B1:Y1
                        com.casapazmino.microservicio_reportes.util.UtilExcel.establecerTexto(hoja, 0, 1,
                                        com.casapazmino.microservicio_reportes.util.UtilExcel
                                                        .aMayusculasSeguras(request.getEmpresa()),
                                        estTitulo);

                        com.casapazmino.microservicio_reportes.util.UtilExcel.combinarCeldas(hoja, 1, 1, 1, 24); // B2:Y2
                        com.casapazmino.microservicio_reportes.util.UtilExcel.establecerTexto(hoja, 1, 1,
                                        "LISTA DE RÉGIMEN LABORAL", estTitulo);

                        // 4) Encabezados (fila visual 6 → índice 5)
                        final int filaEnc = 5;
                        String[] headers = {
                                        "ITEM", "CÓDIGO", "RÉGIMEN", "PAÍS", "CONTINUIDAD LABORAL",
                                        "ANTIGÜEDAD LABORAL",
                                        "PERIODO LABORAL", "DÍAS POR MES", "TRABAJO MÍNIMO (MES)",
                                        "TRABAJO MÍNIMO (HORAS)",
                                        "DÍAS HÁBILES", "DÍAS LIBRES", "DÍAS CALENDARIO", "ACUMULA VACACIONES",
                                        "MÁXIMO DÍAS ACUMULABLES", "VACACIONES POR PERÍODOS", "DETALLE PERÍODOS",
                                        "VACACIONES HÁBILES MES", "VACACIONES CALENDARIO MES", "VACACIONES HÁBILES DÍA",
                                        "VACACIONES CALENDARIO DÍA", "TIPO ANTIGÜEDAD", "AÑOS ANTIGÜEDAD",
                                        "DÍAS ADICIONALES",
                                        "DETALLE RANGOS VARIABLE"
                        };
                        for (int c = 0; c < headers.length; c++) {
                                com.casapazmino.microservicio_reportes.util.UtilExcel.establecerTexto(hoja, filaEnc, c,
                                                headers[c], estEnc);
                        }
                        org.apache.poi.ss.usermodel.Row filaEncRow = com.casapazmino.microservicio_reportes.util.UtilExcel
                                        .asegurarFila(hoja, filaEnc);
                        com.casapazmino.microservicio_reportes.util.UtilExcel.aplicarEstiloAFila(filaEncRow,
                                        headers.length, estEnc);

                        // 5) Datos
                        int fila = filaEnc + 1;
                        int item = 1;
                        if (request.getRegimenes() != null) {
                                for (RegimenDTO r : request.getRegimenes()) {
                                        String textoPeriodos = construirTextoPeriodos(r);
                                        String textoRangos = construirTextoRangos(r);
                                        String tipoAntig = tipoAntiguedad(r);

                                        org.apache.poi.ss.usermodel.Row row = com.casapazmino.microservicio_reportes.util.UtilExcel
                                                        .asegurarFila(hoja, fila++);

                                        set(row, 0, item++, estCentro);
                                        set(row, 1, r.getId(), estCentro);
                                        set(row, 2, nz(r.getDescripcion()), estIzquierda);
                                        set(row, 3, nz(r.getPais()), estIzquierda);
                                        set(row, 4, siNo(bool(r.getContinuidad_laboral())), estCentro);
                                        set(row, 5, siNo(bool(r.getAntiguedad())), estCentro);
                                        set(row, 6, str(r.getMes_periodo()), estIzquierda);
                                        set(row, 7, r.getDias_mes(), estCentro);
                                        set(row, 8, r.getTrabajo_minimo_mes(), estCentro);
                                        set(row, 9, r.getTrabajo_minimo_horas(), estCentro);
                                        set(row, 10, r.getVacacion_dias_laboral(), estCentro);
                                        set(row, 11, r.getVacacion_dias_libre(), estCentro);
                                        set(row, 12, r.getVacacion_dias_calendario(), estCentro);
                                        set(row, 13, siNo(bool(r.getAcumular())), estCentro);
                                        set(row, 14, r.getDias_maximo_acumulacion(), estCentro);
                                        set(row, 15, siNo(bool(r.getVacacion_divisible())), estCentro);
                                        set(row, 16, textoPeriodos, estIzquierda);
                                        set(row, 17, r.getVacacion_dias_laboral_mes(), estCentro);
                                        set(row, 18, r.getVacacion_dias_calendario_mes(), estCentro);
                                        set(row, 19, r.getLaboral_dias(), estCentro);
                                        set(row, 20, r.getCalendario_dias(), estCentro);
                                        set(row, 21, tipoAntig, estCentro);
                                        set(row, 22, r.getAnio_antiguedad(), estCentro);
                                        set(row, 23, r.getDias_antiguedad(), estCentro);
                                        set(row, 24, textoRangos, estIzquierda);
                                }
                        }

                        // 6) Tabla visual (zebra)
                        int ultimaFila = Math.max(fila - 1, filaEnc);
                        com.casapazmino.microservicio_reportes.util.UtilExcel.crearTablaEstilizada(
                                        hoja, "RegimenTabla", filaEnc, 0, ultimaFila, headers.length - 1, true, null);

                        // 7) Anchos (25 columnas)
                        com.casapazmino.microservicio_reportes.util.UtilExcel.establecerAnchosColumnas(hoja, new int[] {
                                        7, 8, 20, 10, 25, 25, 17, 15, 25, 25, 15, 15, 20, 25, 30, 30, 50, 27, 30, 27,
                                        30, 20, 20, 20, 55
                        });

                        wb.write(out);
                        return out.toByteArray();

                } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                }
        }

        private void set(org.apache.poi.ss.usermodel.Row r, int c, Object v, org.apache.poi.ss.usermodel.CellStyle st) {
                com.casapazmino.microservicio_reportes.util.UtilExcel.establecerValor(r, c, v, st);
        }

        // ======================= CSV =======================
        public byte[] generarReporteRegimenesCSV(ReporteRegimenesRequest request) {
                try {
                        String[] headers = {
                                        "ITEM", "CÓDIGO", "RÉGIMEN", "PAÍS", "CONTINUIDAD LABORAL",
                                        "ANTIGÜEDAD LABORAL",
                                        "PERIODO LABORAL", "DÍAS POR MES", "TRABAJO MÍNIMO (MES)",
                                        "TRABAJO MÍNIMO (HORAS)",
                                        "DÍAS HÁBILES", "DÍAS LIBRES", "DÍAS CALENDARIO", "ACUMULA VACACIONES",
                                        "MÁXIMO DÍAS ACUMULABLES", "VACACIONES POR PERÍODOS", "DETALLE PERÍODOS",
                                        "VACACIONES HÁBILES MES", "VACACIONES CALENDARIO MES", "VACACIONES HÁBILES DÍA",
                                        "VACACIONES CALENDARIO DÍA", "TIPO ANTIGÜEDAD", "AÑOS ANTIGÜEDAD",
                                        "DÍAS ADICIONALES",
                                        "DETALLE RANGOS VARIABLE"
                        };

                        StringBuilder sb = new StringBuilder();
                        // encabezado
                        for (int i = 0; i < headers.length; i++) {
                                sb.append(csv(headers[i]));
                                if (i < headers.length - 1)
                                        sb.append(',');
                        }
                        sb.append('\n');

                        int n = 1;
                        if (request.getRegimenes() != null) {
                                for (RegimenDTO r : request.getRegimenes()) {
                                        String tipoAntig = tipoAntiguedad(r);
                                        String textoPeriodos = construirTextoPeriodos(r);
                                        String textoRangos = construirTextoRangos(r);

                                        String[] row = {
                                                        String.valueOf(n++),
                                                        str(r.getId()),
                                                        nz(r.getDescripcion()),
                                                        nz(r.getPais()),
                                                        siNo(bool(r.getContinuidad_laboral())),
                                                        siNo(bool(r.getAntiguedad())),
                                                        str(r.getMes_periodo()),
                                                        str(r.getDias_mes()),
                                                        str(r.getTrabajo_minimo_mes()),
                                                        str(r.getTrabajo_minimo_horas()),
                                                        str(r.getVacacion_dias_laboral()),
                                                        str(r.getVacacion_dias_libre()),
                                                        str(r.getVacacion_dias_calendario()),
                                                        siNo(bool(r.getAcumular())),
                                                        str(r.getDias_maximo_acumulacion()),
                                                        siNo(bool(r.getVacacion_divisible())),
                                                        textoPeriodos,
                                                        str(r.getVacacion_dias_laboral_mes()),
                                                        str(r.getVacacion_dias_calendario_mes()),
                                                        str(r.getLaboral_dias()),
                                                        str(r.getCalendario_dias()),
                                                        tipoAntig,
                                                        str(r.getAnio_antiguedad()),
                                                        str(r.getDias_antiguedad()),
                                                        textoRangos
                                        };

                                        for (int i = 0; i < row.length; i++) {
                                                sb.append(csv(row[i]));
                                                if (i < row.length - 1)
                                                        sb.append(',');
                                        }
                                        sb.append('\n');
                                }
                        }
                        return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);

                } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                }
        }

        private String csv(String v) {
                if (v == null)
                        return "";
                boolean q = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
                String s = v.replace("\"", "\"\"");
                return q ? "\"" + s + "\"" : s;
        }

        // ======================= XML =======================
        public byte[] generarReporteRegimenesXML(ReporteRegimenesRequest request) {
                try {
                        StringBuilder sb = new StringBuilder();
                        sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
                        sb.append("<Regimen_laboral_listado>\n");

                        if (request.getRegimenes() != null) {
                                for (RegimenDTO r : request.getRegimenes()) {
                                        sb.append("  <regimen>\n");
                                        sb.append("    <regimen_laboral id=\"").append(x(r.getId())).append("\">\n");

                                        tag(sb, "descripcion", r.getDescripcion());
                                        tag(sb, "pais", r.getPais());
                                        tag(sb, "continuidad_laboral", siNo(bool(r.getContinuidad_laboral())));
                                        tag(sb, "antiguedad_laboral", siNo(bool(r.getAntiguedad())));
                                        tag(sb, "meses_periodo", r.getMes_periodo()); // ← nombre de nodo como en front
                                        tag(sb, "dias_mes", r.getDias_mes());
                                        tag(sb, "trabajo_minimo_mes", r.getTrabajo_minimo_mes());
                                        tag(sb, "trabajo_minimo_hora", r.getTrabajo_minimo_horas()); // ← nombre de nodo
                                                                                                     // como en front
                                        tag(sb, "dias_habiles", r.getVacacion_dias_laboral());
                                        tag(sb, "dias_libres", r.getVacacion_dias_libre());
                                        tag(sb, "dias_calendario", r.getVacacion_dias_calendario());
                                        tag(sb, "acumula_vacaciones", siNo(bool(r.getAcumular())));
                                        tag(sb, "max_dias_acumulables", r.getDias_maximo_acumulacion());
                                        tag(sb, "vacaciones_por_periodos", siNo(bool(r.getVacacion_divisible())));
                                        tag(sb, "dias_laborales_ganados_mes", r.getVacacion_dias_laboral_mes());
                                        tag(sb, "dias_calendario_ganados_mes", r.getVacacion_dias_calendario_mes());
                                        tag(sb, "dias_laborales_ganados_dia", r.getLaboral_dias());
                                        tag(sb, "dias_calendario_ganados_dia", r.getCalendario_dias());

                                        String tipoAntig = tipoAntiguedad(r);
                                        tag(sb, "tipo_antiguedad", tipoAntig);
                                        tag(sb, "anios_antiguedad", r.getAnio_antiguedad());
                                        tag(sb, "dias_adicionales", r.getDias_antiguedad());

                                        // detalle_vacaciones_periodos
                                        sb.append("      <detalle_vacaciones_periodos>");
                                        if (!bool(r.getVacacion_divisible())) {
                                                sb.append(x("NO APLICA"));
                                        } else if (r.getPeriodos_vacacionales() == null
                                                        || r.getPeriodos_vacacionales().isEmpty()) {
                                                sb.append(x("NO DEFINIDO"));
                                        } else {
                                                sb.append('\n');
                                                for (PeriodoVacacionalDTO p : r.getPeriodos_vacacionales()) {
                                                        sb.append("        <periodo>\n");
                                                        tag(sb, "descripcion", p.getDescripcion(), 10);
                                                        tag(sb, "dias", p.getDias_vacacion(), 10);
                                                        sb.append("        </periodo>\n");
                                                }
                                                sb.append("      ");
                                        }
                                        sb.append("</detalle_vacaciones_periodos>\n");

                                        // detalle_rangos_antiguedad_variable
                                        sb.append("      <detalle_rangos_antiguedad_variable>");
                                        if (!bool(r.getAntiguedad_variable())) {
                                                sb.append(x("NO APLICA"));
                                        } else if (r.getRangos_antiguedad() == null
                                                        || r.getRangos_antiguedad().isEmpty()) {
                                                sb.append(x("NO DEFINIDO"));
                                        } else {
                                                sb.append('\n');
                                                for (RangoAntiguedadDTO g : r.getRangos_antiguedad()) {
                                                        sb.append("        <rango>\n");
                                                        tag(sb, "desde", g.getAnio_desde(), 10);
                                                        tag(sb, "hasta", g.getAnio_hasta(), 10);
                                                        tag(sb, "dias", g.getDias_antiguedad(), 10);
                                                        sb.append("        </rango>\n");
                                                }
                                                sb.append("      ");
                                        }
                                        sb.append("</detalle_rangos_antiguedad_variable>\n");

                                        sb.append("    </regimen_laboral>\n");
                                        sb.append("  </regimen>\n");
                                }
                        }

                        sb.append("</Regimen_laboral_listado>\n");
                        return sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8);

                } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                }
        }

        // ======================= Helpers comunes =======================
        private String nz(String s) {
                return s == null ? "" : s;
        }

        private String str(Object o) {
                return o == null ? "" : String.valueOf(o);
        }

        private boolean bool(Boolean b) {
                return Boolean.TRUE.equals(b);
        }

        private String siNo(boolean b) {
                return b ? "SI" : "NO";
        } // ← sin tilde, igual que en front

        private String tipoAntiguedad(RegimenDTO r) {
                if (bool(r.getAntiguedad_fija()))
                        return "FIJA";
                if (bool(r.getAntiguedad_variable()))
                        return "VARIABLE";
                return "NO APLICA";
        }

        private String construirTextoPeriodos(RegimenDTO r) {
                if (!bool(r.getVacacion_divisible()))
                        return "NO APLICA";
                if (r.getPeriodos_vacacionales() == null || r.getPeriodos_vacacionales().isEmpty())
                        return "NO DEFINIDO";
                return r.getPeriodos_vacacionales().stream()
                                .map(p -> nz(p.getDescripcion()) + ": " + str(p.getDias_vacacion()) + " días")
                                .collect(java.util.stream.Collectors.joining(" | "));
        }

        private String construirTextoRangos(RegimenDTO r) {
                if (!bool(r.getAntiguedad_variable()))
                        return "NO APLICA";
                if (r.getRangos_antiguedad() == null || r.getRangos_antiguedad().isEmpty())
                        return "NO DEFINIDO";
                return r.getRangos_antiguedad().stream()
                                .map(g -> "De " + str(g.getAnio_desde()) + " a " + str(g.getAnio_hasta()) + " años: "
                                                + str(g.getDias_antiguedad()) + " días")
                                .collect(java.util.stream.Collectors.joining(" | "));
        }

        private void tag(StringBuilder sb, String name, Object val) {
                tag(sb, name, val, 6);
        }

        private void tag(StringBuilder sb, String name, Object val, int indent) {
                String s = str(val);
                for (int i = 0; i < indent; i++)
                        sb.append(' ');
                sb.append('<').append(name).append('>').append(x(s)).append("</").append(name).append(">\n");
        }

        private String x(Object v) {
                String s = str(v);
                return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                                .replace("\"", "&quot;").replace("'", "&apos;");
        }

}