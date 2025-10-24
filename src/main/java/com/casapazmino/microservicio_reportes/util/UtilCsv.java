package com.casapazmino.microservicio_reportes.util;

public final class UtilCsv {

    private UtilCsv() { /* utility class */ }

    public static String csvEscape(String v) {
        if (v == null) return "";
        boolean mustQuote = v.contains(",") || v.contains("\"") || v.contains("\n") || v.contains("\r");
        String s = v.replace("\"", "\"\"");
        return mustQuote ? "\"" + s + "\"" : s;
    }
}
