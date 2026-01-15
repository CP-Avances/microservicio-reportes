package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;

public interface ReporteHorasExtraUseCase {
    byte[] generarPdf(ReporteHorasExtraRequest request);

    byte[] generarCsv(ReporteHorasExtraRequest request);

    byte[] generarExcel(ReporteHorasExtraRequest request);

    byte[] generarHtml(ReporteHorasExtraRequest request);
}

