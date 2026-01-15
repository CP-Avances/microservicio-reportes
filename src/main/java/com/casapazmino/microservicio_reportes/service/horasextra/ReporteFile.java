package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;

public interface ReporteFile {

    byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request);

}
