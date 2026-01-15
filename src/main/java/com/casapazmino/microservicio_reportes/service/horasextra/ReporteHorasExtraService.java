package com.casapazmino.microservicio_reportes.service.horasextra;

import com.casapazmino.microservicio_reportes.model.HorasExtra.ReporteHorasExtraRequest;
import com.casapazmino.microservicio_reportes.service.horasextra.interfaces.ReporteFile;
import com.casapazmino.microservicio_reportes.service.horasextra.interfaces.ReporteHorasExtraUseCase;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

@Service
public class ReporteHorasExtraService implements ReporteHorasExtraUseCase {

    private final ReporteFile reportePdf;
    private final ReporteFile reporteCsv;
    private final ReporteFile reporteExcel;
    private final ReporteFile reporteHtml;

    public ReporteHorasExtraService(
            @Qualifier("reporteHorasExtraPdf") ReporteFile reportePdf,
            @Qualifier("reporteHorasExtraCsv") ReporteFile reporteCsv,
            @Qualifier("reporteHorasExtraExcel") ReporteFile reporteExcel,
            @Qualifier("reporteHorasExtraHtml") ReporteFile reporteHtml
    ) {
        this.reportePdf = reportePdf;
        this.reporteCsv = reporteCsv;
        this.reporteExcel = reporteExcel;
        this.reporteHtml = reporteHtml;
    }

    @Override
    public byte[] generarPdf(ReporteHorasExtraRequest request) {
        return reportePdf.generarReporteHorasExtra(request);
    }

    @Override
    public byte[] generarCsv(ReporteHorasExtraRequest request) {
        return reporteCsv.generarReporteHorasExtra(request);
    }

    @Override
    public byte[] generarExcel(ReporteHorasExtraRequest request) {
        return reporteExcel.generarReporteHorasExtra(request);
    }

    @Override
    public byte[] generarHtml(ReporteHorasExtraRequest request) {
        return reporteHtml.generarReporteHorasExtra(request);
    }
}
