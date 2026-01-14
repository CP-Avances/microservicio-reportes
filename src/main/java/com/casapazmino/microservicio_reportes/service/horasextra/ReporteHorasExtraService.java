package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

@Service
public class ReporteHorasExtraService implements ReporteHorasExtraUseCase {

  private final ReporteFile reportePdf;
  private final ReporteFile reporteCsv;

  public ReporteHorasExtraService(
      @Qualifier("reporteHoraExtraPdf") ReporteFile reportePdf,
      @Qualifier("reporteHoraExtraCsv") ReporteFile reporteCsv) {
    this.reportePdf = reportePdf;
    this.reporteCsv = reporteCsv;
  }

  @Override
  public byte[] generarPdf(ReporteHorasExtraRequest request) {
    return reportePdf.generarReporteHorasExtra(request);
  }

  @Override
  public byte[] generarCsv(ReporteHorasExtraRequest request) {
    return reporteCsv.generarReporteHorasExtra(request);
  }
}