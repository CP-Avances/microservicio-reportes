package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion.ReporteTiempoAlimentacionRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTiempoAlimentacionService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/tiempo-alimentacion")
public class ReporteTiempoAlimentacionController {

    @Autowired
    private ReporteTiempoAlimentacionService reporteTiempoAlimentacionService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteTiempoAlimentacion(@RequestBody ReporteTiempoAlimentacionRequest request) {
        byte[] pdfBytes = reporteTiempoAlimentacionService.generarReporteTiempoAlimentacionPDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_tiempo_alimentacion.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }
}
