package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.RegimenLaboral.ReporteRegimenesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteRegimenesService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes/regimen")
@CrossOrigin(origins = "*")
public class ReporteRegimenController {

    private final ReporteRegimenesService reporteRegimenService;

    public ReporteRegimenController(ReporteRegimenesService reporteRegimenService) {
        this.reporteRegimenService = reporteRegimenService;
    }

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReportePDF(@RequestBody ReporteRegimenesRequest request) throws Exception {
        byte[] pdf = reporteRegimenService.generarReporteRegimenesPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Regimen_laboral.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdf);
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteRegimenesRequest request) throws Exception {
        byte[] xlsx = reporteRegimenService.generarReporteRegimenesXLSX(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=RegimenEXCEL.xlsx")
                .contentType(
                        MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(xlsx);
    }

    @PostMapping("/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteRegimenesRequest request) throws Exception {
        byte[] csv = reporteRegimenService.generarReporteRegimenesCSV(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=RegimenCSV.csv")
                .contentType(MediaType.parseMediaType("text/csv"))
                .body(csv);
    }

    @PostMapping("/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteRegimenesRequest request) throws Exception {
        byte[] xml = reporteRegimenService.generarReporteRegimenesXML(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Regimen_laboral.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(xml);
    }
}
