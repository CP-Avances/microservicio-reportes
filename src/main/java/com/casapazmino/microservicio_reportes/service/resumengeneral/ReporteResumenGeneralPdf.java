package com.casapazmino.microservicio_reportes.service.resumengeneral;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;
import org.springframework.stereotype.Component;

@Component("reporteResumenGeneralPdf")
public class ReporteResumenGeneralPdf implements ReporteFileResumenGeneral{

    @Override
    public byte[] generarReporteResumenGeneral(ResumenGeneral request) {
        return new byte[0];
    }
}
