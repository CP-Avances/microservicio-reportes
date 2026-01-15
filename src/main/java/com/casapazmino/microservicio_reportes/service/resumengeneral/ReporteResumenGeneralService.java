package com.casapazmino.microservicio_reportes.service.resumengeneral;

import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;
import com.casapazmino.microservicio_reportes.service.horasextra.ReporteFile;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

@Service
public class ReporteResumenGeneralService implements ReporteResumenGeneralUseCase {

    private final ReporteFileResumenGeneral reportePdf;

    public ReporteResumenGeneralService(@Qualifier("reporteResumenGeneralPdf") ReporteFileResumenGeneral reportePdf) {
        this.reportePdf = reportePdf;
    }



    @Override
    public byte[] generarPdf(ResumenGeneral request) {
        return this.reportePdf.generarReporteResumenGeneral(request);
    }
}
