package com.casapazmino.microservicio_reportes.service.resumengeneral.interfaces;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;

public interface ReporteFileResumenGeneral {
     byte[] generarReporteResumenGeneral(ResumenGeneral request);
}
