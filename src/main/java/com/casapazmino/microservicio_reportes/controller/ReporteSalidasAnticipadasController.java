package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas.ReporteSalidasAnticipadasRequest;
import com.casapazmino.microservicio_reportes.service.ReporteSalidasAnticipadasService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/salidas-anticipadas")
public class ReporteSalidasAnticipadasController {

    @Autowired
    private ReporteSalidasAnticipadasService reporteSalidasAnticipadasService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteSalidasAnticipadas(@RequestBody ReporteSalidasAnticipadasRequest request) {
        byte[] pdfBytes = reporteSalidasAnticipadasService.generarReportePDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_salidas_anticipadas.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }

    // === XLSX ===
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteSalidasAnticipadasXLSX(@RequestBody ReporteSalidasAnticipadasRequest request) {
        byte[] bin = reporteSalidasAnticipadasService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Reporte_Salidas_Anticipadas.xlsx")
                .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }



}
