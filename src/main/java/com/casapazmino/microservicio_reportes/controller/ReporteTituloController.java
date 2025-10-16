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

    // XLSX
    @PostMapping("/titulos/xlsx")
    public ResponseEntity<byte[]> generarReporteTitulosXLSX(@RequestBody ReporteTitulosRequest request) {
        byte[] bin = reporteService.generarReporteTitulosXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=titulos.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/titulos/csv")
    public ResponseEntity<byte[]> generarReporteTitulosCSV(@RequestBody ReporteTitulosRequest request) {
        byte[] bin = reporteService.generarReporteTitulosCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=titulos.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/titulos/xml")
    public ResponseEntity<byte[]> generarReporteTitulosXML(@RequestBody ReporteTitulosRequest request) {
        byte[] bin = reporteService.generarReporteTitulosXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=titulos.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}

