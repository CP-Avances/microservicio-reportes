package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.EstadoCivil.ReporteEstadosCivilRequest;
import com.casapazmino.microservicio_reportes.service.ReporteEstadoCivilService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteEstadoCivilController {

    @Autowired
    private ReporteEstadoCivilService reporteService;

    @PostMapping("/estado-civil/pdf")
    public ResponseEntity<byte[]> generarReporteEstadosCivil(@RequestBody ReporteEstadosCivilRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity
                .ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=estados_civil.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

     // XLSX -> /api/reportes/estados-civil/xlsx
    @PostMapping("/estados-civil/xlsx")
    public ResponseEntity<byte[]> generarReporteEstadosCivilXLSX(@RequestBody ReporteEstadosCivilRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=estados_civil.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV  -> /api/reportes/estados-civil/csv
    @PostMapping("/estados-civil/csv")
    public ResponseEntity<byte[]> generarReporteEstadosCivilCSV(@RequestBody ReporteEstadosCivilRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=estados_civil.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML  -> /api/reportes/estados-civil/xml
    @PostMapping("/estados-civil/xml")
    public ResponseEntity<byte[]> generarReporteEstadosCivilXML(@RequestBody ReporteEstadosCivilRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=estados_civil.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
