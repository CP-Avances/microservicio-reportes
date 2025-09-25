package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Ciudad.ReporteCiudadesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteCiudadesService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteCiudadesController {

    @Autowired
    private ReporteCiudadesService reporteService;

    @PostMapping("/ciudades/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteCiudadesRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ciudades.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    // XLSX
    @PostMapping("/ciudades/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteCiudadesRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ciudades.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/ciudades/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteCiudadesRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ciudades.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/ciudades/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteCiudadesRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ciudades.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }

}
