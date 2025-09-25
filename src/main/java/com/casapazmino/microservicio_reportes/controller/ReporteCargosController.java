package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Cargo.ReporteCargosRequest;
import com.casapazmino.microservicio_reportes.service.ReporteCargosService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteCargosController {

    @Autowired
    private ReporteCargosService reporteService;

    @PostMapping("/cargos/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteCargosRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=tipo_cargos.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

     // XLSX
    @PostMapping("/cargos/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteCargosRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=tipo_cargos.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/cargos/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteCargosRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=tipo_cargos.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/cargos/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteCargosRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=tipo_cargos.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }
}
