package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;

public class ReporteCsv implements ReporteFile {
    @Override
    public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request) {
        return new byte[0];
    }
}
