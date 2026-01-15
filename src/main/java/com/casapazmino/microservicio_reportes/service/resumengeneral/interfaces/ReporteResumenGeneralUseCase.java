package com.casapazmino.microservicio_reportes.service.resumengeneral.interfaces;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;

public interface ReporteResumenGeneralUseCase {
    byte[] generarPdf(ResumenGeneral request);
    byte[] generarCsv(ResumenGeneral request);
    byte[] generarExcel(ResumenGeneral request);
    byte[] generarHtml(ResumenGeneral request);
}
