package com.casapazmino.microservicio_reportes.service.resumengeneral;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;
import com.casapazmino.microservicio_reportes.service.resumengeneral.interfaces.ReporteFileResumenGeneral;
import org.springframework.stereotype.Component;

@Component("reporteResumenGeneralExcel")
public class ReporteResumenExcel implements ReporteFileResumenGeneral {
    @Override
    public byte[] generarReporteResumenGeneral(ResumenGeneral request) {
        return new byte[0];
    }
}
