package com.casapazmino.microservicio_reportes.util;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.xssf.usermodel.*;
import java.util.Base64;
import java.util.regex.Pattern;

/**
 * Utilitarios reutilizables para construir reportes Excel (XLSX) con Apache
 * POI.
 */
public final class UtilExcel {

    private UtilExcel() {
    }

    private static final Pattern PREFIJO_DATA_URI = Pattern.compile("^data:image/[^;]+;base64,",
            Pattern.CASE_INSENSITIVE);

    // -------------------- Imagen / Base64 --------------------

    /** Decodifica imagen base64 con o sin prefijo data:image/...;base64,. */
    public static byte[] decodificarImagenBase64(String base64) {
        if (base64 == null || base64.isEmpty())
            return null;
        String limpio = PREFIJO_DATA_URI.matcher(base64).replaceFirst("");
        try {
            return Base64.getDecoder().decode(limpio);
        } catch (IllegalArgumentException e) {
            return null;
        }
    }

    /**
     * Inserta imagen PNG en (col,fila) con tamaño en píxeles (MOVE_DONT_RESIZE).
     */
    // Inserta el logo ocupando A1:B5 (columnas A-B y filas 1..5), y se redimensiona
    // si cambian anchos/altos.
// En UtilExcel.java
public static void insertarLogoEstandar(Workbook wb, Sheet hoja, byte[] imagenBytes) {
    if (imagenBytes == null || imagenBytes.length == 0) return;

    int tipo = Workbook.PICTURE_TYPE_PNG; // o JPEG si aplica
    int idx = wb.addPicture(imagenBytes, tipo);

    Drawing<?> dibujo = hoja.createDrawingPatriarch();
    CreationHelper helper = wb.getCreationHelper();
    ClientAnchor ancla = helper.createClientAnchor();

    // A1..B5  => col1=0,row1=0 ; col2=2,row2=5
    ancla.setCol1(0);
    ancla.setRow1(0);
    ancla.setCol2(2);
    ancla.setRow2(5);
    ancla.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);

    dibujo.createPicture(ancla, idx);
}


    // -------------------- Celdas / Escritura --------------------

    //Asegura que la fila exista en la hoja antes de escribir en ella.
    public static Row asegurarFila(Sheet hoja, int indiceFila) {
        Row r = hoja.getRow(indiceFila);
        return (r != null) ? r : hoja.createRow(indiceFila);
    }

    //ESTE METODO ESCRIBE TEXTO EN UNA HOJA, INDICANDO LA POSICION CON COORDENADAS DE FILA Y COLUMNA
    public static void establecerTexto(Sheet hoja, int fila, int col, String valor, CellStyle estilo) {
        Row r = asegurarFila(hoja, fila);
        establecerTexto(r, col, valor, estilo);
    }

    //
    public static void establecerTexto(Row fila, int col, String valor, CellStyle estilo) {
        Cell c = fila.getCell(col);
        if (c == null)
            c = fila.createCell(col);
        c.setCellValue(valor == null ? "" : valor);
        if (estilo != null)
            c.setCellStyle(estilo);
    }

    public static void establecerValor(Row fila, int col, Object valor, CellStyle estilo) {
        Cell c = fila.getCell(col);
        if (c == null)
            c = fila.createCell(col);
        if (valor == null)
            c.setBlank();
        else if (valor instanceof Number)
            c.setCellValue(((Number) valor).doubleValue());
        else
            c.setCellValue(String.valueOf(valor));
        if (estilo != null)
            c.setCellStyle(estilo);
    }

    public static void aplicarEstiloAFila(Row fila, int columnas, CellStyle estilo) {
        for (int c = 0; c < columnas; c++) {
            Cell cell = fila.getCell(c);
            if (cell == null)
                cell = fila.createCell(c);
            cell.setCellStyle(estilo);
        }
    }

    // -------------------- Bordes / Estilos masivos --------------------

    public static void agregarBordesDelgados(CellStyle estilo) {
        estilo.setBorderTop(BorderStyle.THIN);
        estilo.setBorderBottom(BorderStyle.THIN);
        estilo.setBorderLeft(BorderStyle.THIN);
        estilo.setBorderRight(BorderStyle.THIN);
    }

    public static void aplicarEstiloARegion(Sheet hoja, int filaInicio, int filaFin,
            int colInicio, int colFin,
            CellStyle estilo, boolean crearSiNoExiste) {
        for (int r = filaInicio; r <= filaFin; r++) {
            Row fila = crearSiNoExiste ? asegurarFila(hoja, r) : hoja.getRow(r);
            if (fila == null)
                continue;
            for (int c = colInicio; c <= colFin; c++) {
                Cell celda = fila.getCell(c);
                if (celda == null && crearSiNoExiste)
                    celda = fila.createCell(c);
                if (celda != null)
                    celda.setCellStyle(estilo);
            }
        }
    }

    public static void establecerAnchosColumnas(Sheet hoja, int[] anchosCaracteres) {
        for (int i = 0; i < anchosCaracteres.length; i++) {
            int w = Math.max(1, anchosCaracteres[i]);
            hoja.setColumnWidth(i, w * 256);
        }
    }

    // -------------------- Merges --------------------

    public static void combinarCeldas(Sheet hoja, int filaIni, int filaFin, int colIni, int colFin) {
        hoja.addMergedRegion(new CellRangeAddress(filaIni, filaFin, colIni, colFin));
    }

    // -------------------- Tabla (XSSFTable) --------------------
    public static void crearTablaEstilizada(XSSFSheet hoja,
            String nombreTabla,
            int filaIni, int colIni,
            int filaFin, int colFin,
            boolean mostrarRayadoFilas,
            boolean[] filtroPorColumna /* no usado en esta versión */) {

        if (filaFin < filaIni)
            filaFin = filaIni; // asegura al menos encabezado

        AreaReference area = new AreaReference(
                new CellReference(filaIni, colIni),
                new CellReference(filaFin, colFin),
                SpreadsheetVersion.EXCEL2007);

        // Crea tabla y configura el CTTable base
        XSSFTable tabla = hoja.createTable(area);
        tabla.setName(nombreTabla);
        tabla.setDisplayName(nombreTabla);

        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable ct = tabla.getCTTable();
        ct.setRef(area.formatAsString());
        ct.setDisplayName(nombreTabla);
        ct.setName(nombreTabla);

        // AutoFilter sobre el rango completo
        if (ct.getAutoFilter() != null) {
            ct.unsetAutoFilter();
        }
        ct.addNewAutoFilter();

        // Estilo de tabla (TableStyleMedium16) + zebra (sin isSet*, con null-check)
        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo tsi = (ct.getTableStyleInfo() != null)
                ? ct.getTableStyleInfo()
                : ct.addNewTableStyleInfo();
        tsi.setName(ConfiguracionExcel.NOMBRE_ESTILO_TABLA); // "TableStyleMedium16"
        tsi.setShowColumnStripes(false);
        tsi.setShowRowStripes(mostrarRayadoFilas);

        // Definir columnas CT (requerido por Excel) con null-checks
        int numColumnas = colFin - colIni + 1;

        org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns ctCols = (ct.getTableColumns() != null)
                ? ct.getTableColumns()
                : ct.addNewTableColumns();
        ctCols.setCount(numColumnas);

        // Asegurar que existan exactamente numColumnas CTTableColumn
        int existentes = ctCols.sizeOfTableColumnArray(); // disponible en todas las versiones recientes
        for (int i = existentes; i < numColumnas; i++) {
            org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn ctc = ctCols.addNewTableColumn();
            ctc.setId(i + 1); // IDs 1..N
            // ctc.setName("Column" + (i+1)); // opcional (Excel autogenera)
        }

        // Nota: omito setFilterButton por columna porque tu esquema no lo expone.
        // Resultado: ícono de filtro visible en todas las columnas (ok para
        // funcionalidad).
    }

    // -------------------- Utilitarios simples --------------------

    public static String aMayusculasSeguras(String s) {
        return s == null ? "" : s.toUpperCase();
    }

    public static String nuloComoVacio(String s) {
        return s == null ? "" : s;
    }
}
