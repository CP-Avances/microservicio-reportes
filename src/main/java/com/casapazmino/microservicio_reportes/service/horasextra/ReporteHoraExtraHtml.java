package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.service.horasextra.interfaces.ReporteFile;

import org.springframework.stereotype.Component;

@Component("reporteHorasExtraHtml")
public class ReporteHoraExtraHtml implements ReporteFile {
    @Override
    public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request) {
        return new byte[0];
    }
}
