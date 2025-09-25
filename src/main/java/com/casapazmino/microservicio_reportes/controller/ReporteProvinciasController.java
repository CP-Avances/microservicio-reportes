package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Provincia.ReporteProvinciasRequest;
import com.casapazmino.microservicio_reportes.service.ReporteProvinciasService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteProvinciasController {

    @Autowired
    private ReporteProvinciasService reporteService;

    @PostMapping("/provincias/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteProvinciasRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=provincias.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    // XLSX
    @PostMapping("/provincias/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteProvinciasRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=provincias.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/provincias/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteProvinciasRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=provincias.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/provincias/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteProvinciasRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=provincias.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
