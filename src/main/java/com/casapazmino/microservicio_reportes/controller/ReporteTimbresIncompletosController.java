package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos.ReporteTimbresIncompletosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTimbresIncompletosService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/timbres-incompletos")
public class ReporteTimbresIncompletosController {

    @Autowired
    private ReporteTimbresIncompletosService reporteService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteTimbresIncompletosRequest request) {
        try {
            byte[] pdf = reporteService.generarReportePDF(request);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Timbres_incompletos.pdf")
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(pdf);

        } catch (Exception e) {
            return ResponseEntity.status(500).build();
        }
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarExcel(@RequestBody ReporteTimbresIncompletosRequest request) {
        byte[] excelBytes = reporteService.generarReporteTimbresIncompletosExcel(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=timbres_incompletos.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(excelBytes);
    }

}
