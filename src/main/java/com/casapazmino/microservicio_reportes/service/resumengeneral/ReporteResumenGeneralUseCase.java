package com.casapazmino.microservicio_reportes.service.resumengeneral;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;

public interface ReporteResumenGeneralUseCase {
    byte[] generarPdf(ResumenGeneral request);
}
