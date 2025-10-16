package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Discapacidad.ReporteDiscapacidadesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteDiscapacidadService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/discapacidades")
@CrossOrigin(origins = "*")
public class ReporteDiscapacidadController {

    private final ReporteDiscapacidadService reporteService;

    public ReporteDiscapacidadController(ReporteDiscapacidadService reporteService) {
        this.reporteService = reporteService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteDiscapacidades(@RequestBody ReporteDiscapacidadesRequest request) {
        byte[] pdf = reporteService.generarReporteDiscapacidadesPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Discapacidades.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }

    // XLSX
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteDiscapacidadesRequest request) {
        byte[] bin = reporteService.generarReporteDiscapacidadesXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=discapacidades.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteDiscapacidadesRequest request) {
        byte[] bin = reporteService.generarReporteDiscapacidadesCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=discapacidades.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteDiscapacidadesRequest request) {
        byte[] bin = reporteService.generarReporteDiscapacidadesXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=discapacidades.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }

}
