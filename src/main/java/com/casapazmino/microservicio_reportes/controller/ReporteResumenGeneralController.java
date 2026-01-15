package com.casapazmino.microservicio_reportes.controller;


import com.casapazmino.microservicio_reportes.model.ResumenGeneral.ResumenGeneral;
import com.casapazmino.microservicio_reportes.service.resumengeneral.ReporteResumenGeneralUseCase;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/reporte/resumen-general")
public class ReporteResumenGeneralController {
 @Autowired
 private ReporteResumenGeneralUseCase reporteResumenGeneralUseCase;

 @PostMapping(value = "/pdf", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_PDF_VALUE)
 public byte[] generarReporteResumenGeneralPdf(@RequestBody ResumenGeneral request) {
    return reporteResumenGeneralUseCase.generarPdf(request);
 }

}
