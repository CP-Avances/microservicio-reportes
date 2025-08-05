package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado.ReporteTiempoLaboradoRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTiempoLaboradoService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/tiempo-laborado")
public class ReporteTiempoLaboradoController {

    @Autowired
    private ReporteTiempoLaboradoService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteTiempoLaboradoRequest request) {
        byte[] pdfBytes = reporteService.generarReporteTiempoLaboradoPDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "tiempo_laborado.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }
}
