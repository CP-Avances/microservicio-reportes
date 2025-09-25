package com.casapazmino.microservicio_reportes.service;

import com.casapazmino.microservicio_reportes.model.RegimenLaboral.*;
import com.casapazmino.microservicio_reportes.util.ConfiguracionPaginaPDF;
import com.casapazmino.microservicio_reportes.util.ReporteUtil;
import com.lowagie.text.*;
import com.lowagie.text.pdf.*;
import org.springframework.stereotype.Service;
import java.awt.Color;
import java.io.ByteArrayOutputStream;

@Service
public class ReporteRegimenesService {

        // METODO QUE GENERA EL PDF
        public byte[] generarReporteRegimenesPDF(ReporteRegimenesRequest request) {
                try {
                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
                        // TIPO Y TAMAÑO DE LA PAGINA DEL REPORTE
                        Document document = new Document(PageSize.A4.rotate());
                        PdfWriter writer = PdfWriter.getInstance(document, baos);
                        writer.setPageEvent(new ConfiguracionPaginaPDF(
                                        request.getUsuario(),
                                        request.getFraseMarcaAgua(),
                                        request.getColorPrincipal()));

                        document.open();

                        // LOGO DE EMPRESA
                        Image logo = ReporteUtil.obtenerLogo(request.getLogoBase64());
                        if (logo != null) {
                                document.add(logo);
                        }

                        // TITULO DE EMPRESA (EKM. CASA PAZMIÑO S.A.)
                        document.add(ReporteUtil.crearTituloEmpresa(request.getEmpresa()));

                        // TITULO DE REPORTE (EJM. REPORTE DE ATRASOS)
                        document.add(ReporteUtil.crearTituloReporte("RÉGIMEN LABORAL"));

                        // COLORES DE LA EMPRESA USADOS EN EL REPORTE
                        Color colorPrincipal = ReporteUtil.convertirHexAColor(request.getColorPrincipal());
                        Color colorSecundario = ReporteUtil.convertirHexAColor(request.getColorSecundario());
                        Color zebra = ReporteUtil.colorZebraClaro();

                        for (RegimenDTO reg : request.getRegimenes()) {
                                int contador = 0;

                                // TABLA DE ENCABEZADO DEL REGIMEN LABORAL
                                PdfPTable tablaCabecera = new PdfPTable(4);
                                tablaCabecera.setWidthPercentage(100);
                                tablaCabecera.setWidths(new float[] { 2.5f, 3, 2.5f, 2 });

                                tablaCabecera.addCell(
                                                ReporteUtil.celdaInfoMixta("PAÍS: ", reg.getPais(), colorPrincipal));
                                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("RÉGIMEN: ", reg.getDescripcion(),
                                                colorPrincipal));
                                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("CONTINUIDAD LABORAL: ",
                                                (reg.getContinuidad_laboral() ? "SÍ" : "NO"), colorPrincipal));
                                tablaCabecera
                                                .addCell(ReporteUtil.celdaInfoMixta("CÓDIGO: ",
                                                                String.valueOf(reg.getId()), colorPrincipal));
                                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("PERIODO LABORAL: ",
                                                reg.getMes_periodo() + " Meses",
                                                colorPrincipal));
                                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("DÍAS POR MES: ",
                                                String.valueOf(reg.getDias_mes()),
                                                colorPrincipal));
                                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("TIEMPO MÍNIMO: ",
                                                (reg.getTrabajo_minimo_mes() > 0
                                                                ? reg.getTrabajo_minimo_mes() + " Meses"
                                                                : reg.getTrabajo_minimo_horas() + " Horas"),
                                                colorPrincipal));
                                tablaCabecera.addCell(ReporteUtil.celdaInfoMixta("ANTIGÜEDAD LABORAL: ",
                                                (reg.getAntiguedad() ? "SÍ" : "NO"), colorPrincipal));

                                // TABLA CONTENEDORA DE LA TABLA ENCABEZADO( ESA NO TIENE NINGUN BORDE) ESTA
                                // SIRVE PARA EL BORDE EXTERNO
                                PdfPTable tablaContenedora = new PdfPTable(1);
                                tablaContenedora.setWidthPercentage(100);
                                PdfPCell contenedor1 = new PdfPCell(tablaCabecera);
                                contenedor1.setBorder(Rectangle.BOX);
                                tablaContenedora.addCell(contenedor1);
                                tablaContenedora.setSpacingAfter(6f);
                                tablaContenedora.setSpacingBefore(10f);
                                document.add(tablaContenedora);

                                // TABLA 1-CONFIGURACION DE VACACIONES
                                PdfPTable configVac = new PdfPTable(2);
                                configVac.setWidthPercentage(100);
                                configVac.setWidths(new float[] { 2.5f, 1.5f });
                                // CELDA QUE OCUPA DOS COLUMNAS DE LA TABLA(ENCABEZADO DE TABLA 1)
                                configVac.addCell(ReporteUtil.crearCelda("CONFIGURACIÓN DE VACACIONES",
                                                ReporteUtil.fuenteEncabezadoTablaData(),
                                                colorPrincipal, 1, 2));
                                configVac.addCell(ReporteUtil.celdaCentro("DÍAS HÁBILES", colorSecundario));
                                configVac.addCell(
                                                ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_laboral()),
                                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                configVac.addCell(ReporteUtil.celdaCentro("DÍAS LIBRES", colorSecundario));
                                configVac.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getVacacion_dias_libre()),
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                configVac.addCell(ReporteUtil.celdaCentro("DÍAS CALENDARIO", colorSecundario));
                                configVac.addCell(ReporteUtil.celdaCentro(
                                                String.valueOf(reg.getVacacion_dias_calendario()),
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                configVac.addCell(ReporteUtil.celdaCentro("ACUMULA VACACIONES", colorSecundario));
                                configVac.addCell(ReporteUtil.celdaCentro(reg.getAcumular() ? "SÍ" : "NO",
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                if (reg.getAcumular()) {
                                        configVac.addCell(ReporteUtil.celdaCentro("MÁXIMO DÍAS ACUMULABLES",
                                                        colorSecundario));
                                        configVac.addCell(ReporteUtil.celdaCentro(
                                                        String.valueOf(reg.getDias_maximo_acumulacion()),
                                                        (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                }
                                configVac.addCell(ReporteUtil.celdaCentro("VACACIONES POR PERÍODOS", colorSecundario));
                                configVac.addCell(ReporteUtil.celdaCentro(reg.getVacacion_divisible() ? "SÍ" : "NO",
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                if (reg.getVacacion_divisible() && reg.getPeriodos_vacacionales() != null
                                                && !reg.getPeriodos_vacacionales().isEmpty()) {
                                        for (PeriodoVacacionalDTO p : reg.getPeriodos_vacacionales()) {
                                                configVac.addCell(ReporteUtil.celdaCentro(p.getDescripcion(),
                                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                                configVac.addCell(ReporteUtil.celdaCentro(
                                                                p.getDias_vacacion() + " días",
                                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                        }
                                }

                                // TABLA 2-VACACIONES GANADAS
                                PdfPTable vacGanadas = new PdfPTable(2);
                                vacGanadas.setWidthPercentage(100);
                                vacGanadas.setWidths(new float[] { 2.5f, 1.5f });
                                // CELDA QUE OCUPA DOS COLUMNAS DE LA TABLA(ENCABEZADO DE TABLA 2)
                                vacGanadas.addCell(ReporteUtil.crearCelda("VACACIONES GANADAS",
                                                ReporteUtil.fuenteEncabezadoTablaData(),
                                                colorPrincipal, 1, 2));
                                vacGanadas.addCell(ReporteUtil.celdaCentro("POR MES (HÁBILES)", colorSecundario));
                                vacGanadas.addCell(ReporteUtil.celdaCentro(
                                                String.valueOf(reg.getVacacion_dias_laboral_mes()),
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                vacGanadas.addCell(ReporteUtil.celdaCentro("POR MES (CALENDARIO)", colorSecundario));
                                vacGanadas.addCell(ReporteUtil.celdaCentro(
                                                String.valueOf(reg.getVacacion_dias_calendario_mes()),
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                vacGanadas.addCell(ReporteUtil.celdaCentro("POR DÍA (HÁBILES)", colorSecundario));
                                vacGanadas.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getLaboral_dias()),
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                vacGanadas.addCell(ReporteUtil.celdaCentro("POR DÍA (CALENDARIO)", colorSecundario));
                                vacGanadas.addCell(ReporteUtil.celdaCentro(String.valueOf(reg.getCalendario_dias()),
                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));

                                // TABLA 3-CONFIGURACION DE ANTIGUEDAD
                                PdfPTable antiguedad = new PdfPTable(2);
                                antiguedad.setWidthPercentage(100);
                                antiguedad.setWidths(new float[] { 2.5f, 1.5f });
                                // CELDA QUE OCUPA DOS COLUMNAS DE LA TABLA(ENCABEZADO DE TABLA 3)
                                antiguedad.addCell(ReporteUtil.crearCelda("CONFIGURACIÓN DE ANTIGÜEDAD",
                                                ReporteUtil.fuenteEncabezadoTablaData(),
                                                colorPrincipal, 1, 2));
                                if (Boolean.TRUE.equals(reg.getAntiguedad())) {
                                        if (Boolean.TRUE.equals(reg.getAntiguedad_fija())) {
                                                antiguedad.addCell(ReporteUtil.celdaCentro("TIPO", colorSecundario));
                                                antiguedad
                                                                .addCell(ReporteUtil.celdaCentro("FIJA",
                                                                                (contador++ % 2 == 0) ? zebra
                                                                                                : Color.WHITE));
                                                antiguedad.addCell(ReporteUtil.celdaCentro("AÑOS ANTIGÜEDAD",
                                                                colorSecundario));
                                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                                                String.valueOf(reg.getAnio_antiguedad()),
                                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                                antiguedad.addCell(ReporteUtil.celdaCentro("DÍAS ADICIONALES",
                                                                colorSecundario));
                                                antiguedad.addCell(ReporteUtil.celdaCentro(
                                                                String.valueOf(reg.getDias_antiguedad()),
                                                                (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                        } else if (Boolean.TRUE.equals(reg.getAntiguedad_variable())
                                                        && reg.getRangos_antiguedad() != null
                                                        && !reg.getRangos_antiguedad().isEmpty()) {
                                                antiguedad.addCell(ReporteUtil.celdaCentro("TIPO", colorSecundario));
                                                antiguedad.addCell(
                                                                ReporteUtil.celdaCentro("VARIABLE",
                                                                                (contador++ % 2 == 0) ? zebra
                                                                                                : Color.WHITE));
                                                for (RangoAntiguedadDTO r : reg.getRangos_antiguedad()) {
                                                        antiguedad.addCell(ReporteUtil.celdaCentro(
                                                                        "Desde " + r.getAnio_desde() + " hasta "
                                                                                        + r.getAnio_hasta() + " años",
                                                                        colorSecundario));
                                                        antiguedad.addCell(ReporteUtil.celdaCentro(
                                                                        r.getDias_antiguedad() + " días",
                                                                        (contador++ % 2 == 0) ? zebra : Color.WHITE));
                                                }
                                        }
                                } else {
                                        PdfPCell celdaNoAplica = ReporteUtil.celdaCentro("NO APLICA",
                                                        (contador++ % 2 == 0) ? zebra : Color.WHITE);
                                        celdaNoAplica.setColspan(2);
                                        antiguedad.addCell(celdaNoAplica);

                                }

                                int filasVac = configVac.size();
                                int filasGanadas = vacGanadas.size();
                                int filasAntig = antiguedad.size();
                                int maxFilas = Math.max(filasVac, Math.max(filasGanadas, filasAntig));

                                // SE RELLENA DINAMICAMENTE LAS TABLAS QUE OCUPAN MENOS FILAS DEPENDIENDO LA
                                // INFORMACION
                                while (configVac.size() < maxFilas) {
                                        PdfPCell celda1 = ReporteUtil.celdaCentro("", Color.WHITE);
                                        celda1.setBorder(Rectangle.NO_BORDER);
                                        PdfPCell celda2 = ReporteUtil.celdaCentro("", Color.WHITE);
                                        celda2.setBorder(Rectangle.NO_BORDER);
                                        configVac.addCell(celda1);
                                        configVac.addCell(celda2);
                                }
                                while (vacGanadas.size() < maxFilas) {
                                        PdfPCell celda1 = ReporteUtil.celdaCentro("", Color.WHITE);
                                        celda1.setBorder(Rectangle.NO_BORDER);
                                        PdfPCell celda2 = ReporteUtil.celdaCentro("", Color.WHITE);
                                        celda2.setBorder(Rectangle.NO_BORDER);
                                        vacGanadas.addCell(celda1);
                                        vacGanadas.addCell(celda2);
                                }
                                while (antiguedad.size() < maxFilas) {
                                        PdfPCell celda1 = ReporteUtil.celdaCentro("", Color.WHITE);
                                        celda1.setBorder(Rectangle.NO_BORDER);
                                        PdfPCell celda2 = ReporteUtil.celdaCentro("", Color.WHITE);
                                        celda2.setBorder(Rectangle.NO_BORDER);
                                        antiguedad.addCell(celda1);
                                        antiguedad.addCell(celda2);
                                }

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
                                contenedor.setWidthPercentage(100);
                                contenedor.setWidths(new float[] { 3f, 3f, 3f });
                                contenedor.addCell(celdaConfigVac);
                                contenedor.addCell(celdaVacGanadas);
                                contenedor.addCell(celdaAntiguedad);
                                contenedor.setSpacingAfter(10f);

                                contador++;
                                document.add(contenedor);
                        }
                        document.close();
                        return baos.toByteArray();
                } catch (Exception e) {
                        e.printStackTrace();
                        return null;
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