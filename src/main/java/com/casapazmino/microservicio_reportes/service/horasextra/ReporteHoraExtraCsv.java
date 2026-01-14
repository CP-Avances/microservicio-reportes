package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import org.springframework.stereotype.Component;

@Component("reporteHoraExtraCsv")
public class ReporteHoraExtraCsv implements ReporteFile {

  @Override
  public byte[] generarReporteHorasExtra(ReporteHorasExtraRequest request) {
    return new byte[0];
  }
}
