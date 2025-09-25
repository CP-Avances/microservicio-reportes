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

    /** Estilo de título general: negrita, tamaño 14, centrado. */
    public static CellStyle crearEstiloTitulo(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        Font f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints(TAMANIO_FUENTE_TITULO);
        cs.setFont(f);
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
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
        return cs;
    }

    /** Estilo de celdas centradas con borde delgado. */
    public static CellStyle crearEstiloCentroConBorde(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(HorizontalAlignment.CENTER);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }

    /** Estilo de celdas a la izquierda con borde delgado. */
    public static CellStyle crearEstiloIzquierdaConBorde(XSSFWorkbook wb) {
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(HorizontalAlignment.LEFT);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        UtilExcel.agregarBordesDelgados(cs);
        return cs;
    }
}
