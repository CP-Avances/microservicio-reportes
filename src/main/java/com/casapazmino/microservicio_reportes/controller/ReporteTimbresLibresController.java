package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.TimbresLibres.ReporteTimbresLibresRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTimbresLibresService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/timbres-libres")
@CrossOrigin(origins = "*")
public class ReporteTimbresLibresController {

    @Autowired
    private ReporteTimbresLibresService service;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPdf(@RequestBody ReporteTimbresLibresRequest request) {
        byte[] pdf = service.generarReportePDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        String nombre = "TimbresLibres.pdf";
        headers.setContentDispositionFormData("attachment", nombre);

        return ResponseEntity.ok().headers(headers).body(pdf);
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarXlsx(@RequestBody ReporteTimbresLibresRequest request) {
        byte[] xlsx = service.generarReporteXLSX(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        String nombre = "TimbresLibres.xlsx";
        headers.setContentDispositionFormData("attachment", nombre);

        return ResponseEntity.ok().headers(headers).body(xlsx);
    }
}
