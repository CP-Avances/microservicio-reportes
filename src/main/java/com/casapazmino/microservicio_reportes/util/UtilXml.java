package com.casapazmino.microservicio_reportes.util;

public final class UtilXml {

    private UtilXml() {}

    /** Escape XML básico para texto y atributos (null-safe). */
    public static String xmlEsc(Object v) {
        String s = (v == null) ? "" : String.valueOf(v);
        // Importante: primero ampersand, luego el resto
        return s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }
}
