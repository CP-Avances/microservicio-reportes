package com.casapazmino.microservicio_reportes.service.resumengeneral;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;
import com.casapazmino.microservicio_reportes.service.resumengeneral.interfaces.ReporteFileResumenGeneral;
import com.casapazmino.microservicio_reportes.service.resumengeneral.interfaces.ReporteResumenGeneralUseCase;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

@Service
public class ReporteResumenGeneralService implements ReporteResumenGeneralUseCase {

    private final ReporteFileResumenGeneral reportePdf;
    private final ReporteFileResumenGeneral reporteCsv;
    private final ReporteFileResumenGeneral reporteExcel;
    private final ReporteFileResumenGeneral reporteHtml;

    public ReporteResumenGeneralService(
            @Qualifier("reporteResumenGeneralPdf") ReporteFileResumenGeneral reportePdf,
            @Qualifier("reporteResumenGeneralCsv") ReporteFileResumenGeneral reporteCsv,
            @Qualifier("reporteResumenGeneralExcel") ReporteFileResumenGeneral reporteExcel,
            @Qualifier("reporteResumenGeneralHtml") ReporteFileResumenGeneral reporteHtml
    ) {
        this.reportePdf = reportePdf;
        this.reporteCsv = reporteCsv;
        this.reporteExcel = reporteExcel;
        this.reporteHtml = reporteHtml;
    }

    @Override
    public byte[] generarPdf(ResumenGeneral request) {
        return this.reportePdf.generarReporteResumenGeneral(request);
    }

    @Override
    public byte[] generarCsv(ResumenGeneral request) {
        return this.reporteCsv.generarReporteResumenGeneral(request);
    }

    @Override
    public byte[] generarExcel(ResumenGeneral request) {
        return this.reporteExcel.generarReporteResumenGeneral(request);
    }

    @Override
    public byte[] generarHtml(ResumenGeneral request) {
        return this.reporteHtml.generarReporteResumenGeneral(request);
    }
}

