package com.casapazmino.microservicio_reportes.util;

import org.openpdf.text.*;
import org.openpdf.text.pdf.*;
import java.awt.Color;
import java.text.SimpleDateFormat;
import java.util.Base64;
import java.util.Date;
import java.util.Locale;

public class ReporteUtil {

    /*
     * Resumen de refactor de estilos PDF:
     * - Nuevos helpers centralizados: fuenteTituloEmpresa(), fuenteSubtitulo(),
     *   crearSubtituloPeriodo(), celdaEncabezadoTabla(...),
     *   celdaDataCentro(...), celdaDataIzquierda(...).
     * - Unificaciones: celdaEncabezado(...) ahora delega al encabezado estándar y
     *   las celdas centradas/izquierda reutilizan las celdas de datos.
     * - Reportes ajustados: servicios de catálogos y asistencia (Cargos,
     *   Ciudades, Coordenadas, Departamentos, Discapacidad, Empleado,
     *   EstadoCivil, Feriados, Género, Horarios, ModalidadLaboral,
     *   Nacionalidades, NivelTitulo, Parámetros, Provincias, Regímenes,
     *   Relojes, Roles, Sucursales, Título, Vacuna) consumen ahora el estilo
     *   común para encabezados y filas.
     */

    // Convertir código HEX (como "#E5E7E9") a objeto Color
    public static Color convertirHexAColor(String hex) {
        try {
            return Color.decode(hex);
        } catch (Exception e) {
            return Color.LIGHT_GRAY;
        }
    }

    // Aplicar tamaño y alineación al logo
    public static void aplicarEstiloLogo(Image logo) {
        logo.scaleToFit(120, 120);
        logo.setAlignment(Image.LEFT);
    }

    // Fuente de texto general
    public static Font fuenteTexto() {
        return FontFactory.getFont(FontFactory.HELVETICA, 6.5f);
    }

    // Fuente para el nombre de la empresa
    public static Font fuenteTituloEmpresa() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 14);
    }

    // Fuente para el título del reporte
    public static Font fuenteTituloReporte() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 12);
    }

    // Fuente para subtítulos / periodos
    public static Font fuenteSubtitulo() {
        return FontFactory.getFont(FontFactory.HELVETICA, 10);
    }

    // Crear titulo del periodo (alias de subtítulo)
    public static Paragraph crearTituloPeriodo(String texto) {
        return crearSubtituloPeriodo(texto);
    }

    // Crear subtítulo / periodo estándar
    public static Paragraph crearSubtituloPeriodo(String texto) {
        Paragraph p = new Paragraph(texto, fuenteSubtitulo());
        p.setAlignment(Element.ALIGN_CENTER);
        p.setSpacingAfter(10f);
        return p;
    }

    // Color zebra claro reutilizable
    public static Color colorZebraClaro() {
        return new Color(0xE5, 0xE7, 0xE9); // rgb(212, 212, 212)
    }

    // Celda alineada al centro
    public static PdfPCell celdaCentro(String texto, Font fuente) {
        return celdaDataCentro(texto, null, fuente);
    }

    public static PdfPCell celdaCentro(String texto) {
        return celdaDataCentro(texto, null, fuenteTexto());
    }

    // Celda alineada a la izquierda
    public static PdfPCell celdaIzquierda(String texto, Font fuente) {
        return celdaDataIzquierda(texto, null, fuente);
    }

    public static PdfPCell celdaIzquierda(String texto) {
        return celdaDataIzquierda(texto, null, fuenteTexto());
    }

    // Celda gris para info de empleado
    public static PdfPCell celdaInfoEmpleado(String texto) {
        PdfPCell celda = new PdfPCell(new Phrase(texto, fuenteTexto()));
        celda.setBackgroundColor(new Color(227, 227, 227));
        celda.setPadding(5);
        return celda;
    }

    // Celda de encabezado con fondo
    public static PdfPCell celdaEncabezado(String texto, Color fondo) {
        return celdaEncabezadoTabla(texto, fondo);
    }

    // Celda especial para número de día (horario)
    public static PdfPCell celdaEncabezadoDia(int dia) {
        PdfPCell celda = new PdfPCell(new Phrase(String.format("%02d", dia), fuenteEncabezado()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4);
        return celda;
    }

    // Celdas para nomenclatura
    public static PdfPCell celdaNomenclatura(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        celda.setMinimumHeight(20f);
        celda.setNoWrap(true);
        return celda;
    }

    public static PdfPCell celdaNomenclaturaDescripcion(String texto, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        celda.setMinimumHeight(20f);
        celda.setNoWrap(true);
        return celda;
    }

    // FormatearFecha
    public static String formatearFechaConDia(String fechaOriginal) {
        try {
            SimpleDateFormat entrada = new SimpleDateFormat("yyyy-MM-dd");
            SimpleDateFormat salida = new SimpleDateFormat("EEE. dd/MM/yyyy", new Locale("es", "ES"));
            Date fecha = entrada.parse(fechaOriginal);
            String resultado = salida.format(fecha);

            return resultado.substring(0, 1).toUpperCase() + resultado.substring(1);
        } catch (Exception e) {
            return fechaOriginal;
        }
    }

    // Celda izquierda con fondo personalizado
    public static PdfPCell celdaIzquierda(String texto, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuenteTexto()));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setBackgroundColor(fondo);
        celda.setPadding(4f);
        return celda;
    }

    public static String traducirAccion(String codigo) {
        if (codigo == null)
            return "";

        switch (codigo.trim().toUpperCase()) {
            case "E":
                return "Entrada";
            case "S":
                return "Salida";
            case "I/A":
                return "Inicio alimentación";
            case "F/A":
                return "Fin alimentación";
            case "D":
                return "Desconocido";
            default:
                return codigo;
        }
    }

    ///////////////////////////////////////////////////////
    // ORDENAMIENTO DE METODOS PARA REPORTES//
    ///////////////////////////////////////////////////////

    // METODO USADO PARA OBTENER LOGO (Convertir logo base64 a Image)
    public static Image obtenerLogo(String base64String) throws Exception {
        if (base64String != null && base64String.contains("base64,")) {
            String base64 = base64String.split(",")[1];
            byte[] imageBytes = Base64.getDecoder().decode(base64);
            Image logo = Image.getInstance(imageBytes);
            aplicarEstiloLogo(logo);
            return logo;
        }
        return null;
    }

    // METODO USADO PARA CREA TITULO EMPRESA DEL REPORTE (Estilos)
    public static Paragraph crearTituloEmpresa(String texto) {
        Paragraph p = new Paragraph(texto, fuenteTituloEmpresa());
        p.setAlignment(Element.ALIGN_CENTER);
        p.setSpacingBefore(-35f);
        p.setSpacingAfter(2f);
        return p;
    }

    // METODO PARA CREAR TITULO DEL REPORTE(ESTILOS)
    public static Paragraph crearTituloReporte(String texto) {
        Paragraph p = new Paragraph(texto, fuenteTituloReporte());
        p.setAlignment(Element.ALIGN_CENTER);
        p.setSpacingAfter(0f);
        return p;
    }

    // METODO PARA FUENTE DE TABLA ENCABEZADO
    public static Font fuenteEncabezado() {
        return FontFactory.getFont(FontFactory.HELVETICA, 9);
    }

    // METODO PARA FUENTE DE ENCABEZADO DE TABLA DATA (estándar)
    public static Font fuenteEncabezadoTablaData() {
        return FontFactory.getFont(FontFactory.HELVETICA_BOLD, 7);
    }

    // METODO PARA FUENTE DE ENCABEZADO DE TABLA DATA
    public static Font fuenteTablaData() {
        return FontFactory.getFont(FontFactory.HELVETICA, 7);
    }

    // METODO PARA NECABEZADOS QUE TIENEN CABEZERA Y SU INFOMRACION CORRESPONDIENTE
    // (PAIS: ECUADOR)
    public static PdfPCell celdaInfoMixta(String etiqueta, String valor, Color fondo) {
        Phrase contenido = new Phrase();
        contenido.add(new Chunk(etiqueta + " ", ReporteUtil.fuenteEncabezado()));
        contenido.add(new Chunk(valor != null ? valor : "", ReporteUtil.fuenteEncabezado()));

        PdfPCell celda = new PdfPCell(contenido);
        celda.setBackgroundColor(fondo);
        celda.setPadding(2f);
        celda.setBorder(Rectangle.NO_BORDER);
        return celda;
    }

    // METODO CON LA CAPACIDAD DE UNIR VARIAS FILAS Y COLUMNAS
    public static PdfPCell crearCelda(String texto, Font fuente, Color fondo, int rowspan, int colspan) {
        PdfPCell celda = crearCelda(texto, fuente, fondo);
        celda.setRowspan(rowspan);
        celda.setColspan(colspan);
        return celda;
    }

    // METODO PARA CREAR CELDA CON DATOS CENTRADOS A LA CELDA
    public static PdfPCell celdaCentro(String texto, Color fondo) {
        return celdaDataCentro(texto, fondo, fuenteTablaData());
    }

    // Crear celda con texto, fuente y fondo personalizado
    public static PdfPCell crearCelda(String texto, Font fuente, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setBackgroundColor(fondo);
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setPadding(4f);
        return celda;
    }

    // Encabezado de tabla estándar
    public static PdfPCell celdaEncabezadoTabla(String texto, Color fondo) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuenteEncabezadoTablaData()));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        celda.setBackgroundColor(fondo);
        celda.setPadding(5f);
        return celda;
    }

    // Celdas de datos centradas
    public static PdfPCell celdaDataCentro(String texto, Color fondo) {
        return celdaDataCentro(texto, fondo, fuenteTablaData());
    }

    public static PdfPCell celdaDataCentro(String texto, Color fondo, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_CENTER);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        if (fondo != null) {
            celda.setBackgroundColor(fondo);
        }
        celda.setPadding(4f);
        return celda;
    }

    // Celdas de datos alineadas a la izquierda
    public static PdfPCell celdaDataIzquierda(String texto, Color fondo) {
        return celdaDataIzquierda(texto, fondo, fuenteTablaData());
    }

    public static PdfPCell celdaDataIzquierda(String texto, Color fondo, Font fuente) {
        PdfPCell celda = new PdfPCell(new Phrase(texto != null ? texto : "", fuente));
        celda.setHorizontalAlignment(Element.ALIGN_LEFT);
        celda.setVerticalAlignment(Element.ALIGN_MIDDLE);
        if (fondo != null) {
            celda.setBackgroundColor(fondo);
        }
        celda.setPadding(4f);
        return celda;
    }

}
