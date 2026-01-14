package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;

public interface ReporteFile {

  public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request);

}
