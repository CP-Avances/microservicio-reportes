package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Horario.ReporteHorariosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteHorariosService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/horarios")
public class ReporteHorariosController {

    private final ReporteHorariosService reporteHorariosService;

    public ReporteHorariosController(ReporteHorariosService reporteHorariosService) {
        this.reporteHorariosService = reporteHorariosService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarPDF(@RequestBody ReporteHorariosRequest request) {
        byte[] pdfBytes = reporteHorariosService.generarReportePDF(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Horarios.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

        // XLSX
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarXLSX(@RequestBody ReporteHorariosRequest request) {
        byte[] bin = reporteHorariosService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Horarios.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarCSV(@RequestBody ReporteHorariosRequest request) {
        byte[] bin = reporteHorariosService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Horarios.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarXML(@RequestBody ReporteHorariosRequest request) {
        byte[] bin = reporteHorariosService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Horarios.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
