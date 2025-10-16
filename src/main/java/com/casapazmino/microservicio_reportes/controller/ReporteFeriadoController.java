package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Feriado.ReporteFeriadosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteFeriadosService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/feriados")
public class ReporteFeriadoController {

    private final ReporteFeriadosService reporteFeriadosService;

    public ReporteFeriadoController(ReporteFeriadosService reporteFeriadosService) {
        this.reporteFeriadosService = reporteFeriadosService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteFeriadosRequest request) {
        byte[] pdf = reporteFeriadosService.generarReporteFeriadosPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=feriados.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }

    // XLSX
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarXLSX(@RequestBody ReporteFeriadosRequest request) {
        byte[] xlsx = reporteFeriadosService.generarReporteFeriadosXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=FeriadosEXCEL.xlsx")
                .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(xlsx);
    }

    // CSV
    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarCSV(@RequestBody ReporteFeriadosRequest request) {
        byte[] csv = reporteFeriadosService.generarReporteFeriadosCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=feriados.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(csv);
    }

    // XML
    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarXML(@RequestBody ReporteFeriadosRequest request) {
        byte[] xml = reporteFeriadosService.generarReporteFeriadosXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=feriados.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(xml);
    }
}
