package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Titulo.ReporteTitulosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTituloService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteTituloController {

    @Autowired
    private ReporteTituloService reporteService;

    @PostMapping("/titulos/pdf")
    public ResponseEntity<byte[]> generarReporteTitulos(@RequestBody ReporteTitulosRequest request) {
        byte[] pdf = reporteService.generarReporteTitulosPDF(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=titulos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }
}

