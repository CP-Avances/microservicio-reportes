package com.casapazmino.microservicio_reportes.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Estilos y constantes comunes para reportes Excel.
 */
public final class ConfiguracionExcel {

    private ConfiguracionExcel() {}

    public static final String MIME_XLSX =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public static final String NOMBRE_ESTILO_TABLA = "TableStyleMedium16";

    public static final short TAMANIO_FUENTE_TITULO = 14;
    public static final short TAMANIO_FUENTE_ENCABEZADO = 11;
    public static final short TAMANIO_FUENTE_NORMAL = 10;

    // Colores (similares a tu PDF)
    private static final short VERDE_FRANJA = IndexedColors.LIGHT_GREEN.getIndex();   // verde suave
    private static final short GRIS_CAB = IndexedColors.GREY_25_PERCENT.getIndex();  // gris claro
    private static final short AZUL_HEAD = IndexedColors.BLUE_GREY.getIndex();       // azul para header
    private static final short AZUL_SUB  = IndexedColors.PALE_BLUE.getIndex();       // azul claro subheader
    private static final short AZUL_TOTAL = IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex();
    private static final short AZUL_PEND_TIT = IndexedColors.LIGHT_TURQUOISE.getIndex();

    /** Estilo de título general: negrita, tamaño 14, centrado. */
    public static CellStyle crearEstiloTitulo(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_TITULO);
        cs.setFont(f);
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(true);
        return cs;
    }

    /** Estilo de encabezado de tabla: negrita, centrado. */
    public static CellStyle crearEstiloEncabezadoTabla(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_ENCABEZADO);
        cs.setFont(f);
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(true);
        return cs;
    }

    /** Estilo de celdas centradas con borde delgado. */
    public static CellStyle crearEstiloCentroConBorde(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Estilo de celdas a la izquierda con borde delgado. */
    public static CellStyle crearEstiloIzquierdaConBorde(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    // =====================================================================================
    // ✅ NUEVOS ESTILOS PARA LA HOJA "Kardex_Reporte" (igual al PDF / imagen)
    // =====================================================================================

    /** Franja verde (SUCURSAL / DEPARTAMENTO / etc). */
    public static CellStyle crearEstiloFranjaVerde(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_ENCABEZADO);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(VERDE_FRANJA);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Cabecera gris (CIUDAD / C.C. / COD / EMPLEADO). */
    public static CellStyle crearEstiloCabeceraGris(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(GRIS_CAB);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Estilo fila periodo (F.INICIO / F.FIN / etc). */
    public static CellStyle crearEstiloPeriodo(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Header azul (Detalle / Desde / Hasta / Descuento / Saldo). */
    public static CellStyle crearEstiloHeaderAzul(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setColor(IndexedColors.WHITE.getIndex());
        f.setFontHeightInPoints(TAMANIO_FUENTE_ENCABEZADO);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(AZUL_HEAD);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Subheader azul claro (Días/Hor/Min). */
    public static CellStyle crearEstiloSubHeaderAzul(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(AZUL_SUB);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Etiqueta info (opcional si haces "Label: Valor" separado). */
    public static CellStyle crearEstiloInfoLabel(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Valor info (opcional). */
    public static CellStyle crearEstiloInfoValue(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(false);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Título de cajas (Antigüedad / Proporcional / Liquidación). */
    public static CellStyle crearEstiloCajaTitulo(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(AZUL_TOTAL);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Dato dentro de caja (valores de días/hor/min). */
    public static CellStyle crearEstiloCajaDato(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(false);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Título para sección Pendientes (barra celeste). */
    public static CellStyle crearEstiloPendientesTitulo(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(AZUL_PEND_TIT);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Estilo para fila Total estimado. */
    public static CellStyle crearEstiloTotal(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_NORMAL);
        cs.setFont(f);

        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);

        cs.setFillForegroundColor(AZUL_TOTAL);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cs.setWrapText(true);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }
}