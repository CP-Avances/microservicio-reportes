package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Horario.ReporteHorariosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteHorariosService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/horarios")
public class ReporteHorariosController {

    private final ReporteHorariosService reporteHorariosService;

    public ReporteHorariosController(ReporteHorariosService reporteHorariosService) {
        this.reporteHorariosService = reporteHorariosService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteHorariosRequest request) {
        byte[] pdfBytes = reporteHorariosService.generarReportePDF(request);

        if (pdfBytes == null) {
            return ResponseEntity.internalServerError().build();
        }

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=Horarios.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }
}
