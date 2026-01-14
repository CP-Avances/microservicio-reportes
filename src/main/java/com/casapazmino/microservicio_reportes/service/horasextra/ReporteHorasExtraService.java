package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import org.springframework.stereotype.Service;

@Service
public class ReporteHorasExtraService implements ReporteServiceInterface {

  private final ReporteFile reporteFile;

  public ReporteHorasExtraService(ReporteFile reporteFile) {
    this.reporteFile = reporteFile;
  }

  public byte[] generarReporte(ReporteHorasExtraRequest request) {
    return reporteFile.generarReporteHorasExtra(request);
  }
}