package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteFaltas.ReporteFaltasRequest;
import com.casapazmino.microservicio_reportes.service.ReporteFaltasService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/faltas")
public class ReporteFaltasController {

    private final ReporteFaltasService reporteFaltasService;

    public ReporteFaltasController(ReporteFaltasService reporteFaltasService) {
        this.reporteFaltasService = reporteFaltasService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteFaltasRequest request) {
        byte[] pdfBytes = reporteFaltasService.generarReporteFaltasPDF(request);

        if (pdfBytes == null) {
            return ResponseEntity.internalServerError().build();
        }

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=FaltasUsuarios.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
