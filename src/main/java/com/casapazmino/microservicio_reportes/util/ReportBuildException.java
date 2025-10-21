package com.casapazmino.microservicio_reportes.util;

/**
 * Excepción de dominio para fallos internos al construir reportes
 * (PDF/XLSX/CSV/XML). Úsala cuando el proceso de generación falle por
 * causas internas (iText/Lowagie, IO, NPE inesperado, etc.).
 *
 * Objetivo: permitir que el Controller responda 500 de forma uniforme.
 *
 * Ejemplo de uso:
 *   try {
 *     // construir PDF...
 *   } catch (Exception e) {
 *     throw new ReportBuildException("No se pudo generar ResumenAsistencia.pdf", e);
 *   }
 */
public class ReportBuildException extends RuntimeException {
    private static final long serialVersionUID = 1L;

    public ReportBuildException(String message) {
        super(message);
    }

    public ReportBuildException(String message, Throwable cause) {
        super(message, cause);
    }

    public ReportBuildException(Throwable cause) {
        super(cause);
    }
}
