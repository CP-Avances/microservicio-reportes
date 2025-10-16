package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteAtrasos.ReporteAtrasosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteAtrasosService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/atrasos")
public class ReporteAtrasosController {

    @Autowired
    private ReporteAtrasosService reporteAtrasosService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteAtrasos(@RequestBody ReporteAtrasosRequest request) {
        byte[] pdfBytes = reporteAtrasosService.generarReportePDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_atrasos.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteAtrasosExcel(@RequestBody ReporteAtrasosRequest request) {
        byte[] excelBytes = reporteAtrasosService.generarReporteAtrasosExcel(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "reporte_atrasos.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(excelBytes);
    }
}
