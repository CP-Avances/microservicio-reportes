package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Dispositivo.ReporteRelojesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteRelojesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteRelojesController {

    @Autowired
    private ReporteRelojesService reporteService;

    @PostMapping("/relojes/pdf")
    public ResponseEntity<byte[]> generarReporteRelojes(@RequestBody ReporteRelojesRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=relojes.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    // XLSX
    @PostMapping("/relojes/xlsx")
    public ResponseEntity<byte[]> generarReporteRelojesXLSX(@RequestBody ReporteRelojesRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=relojes.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/relojes/csv")
    public ResponseEntity<byte[]> generarReporteRelojesCSV(@RequestBody ReporteRelojesRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=relojes.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/relojes/xml")
    public ResponseEntity<byte[]> generarReporteRelojesXML(@RequestBody ReporteRelojesRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=relojes.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
