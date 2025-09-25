package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Sucursal.ReporteSucursalesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteSucursalesService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteSucursalesController {

    @Autowired
    private ReporteSucursalesService reporteService;

    @PostMapping("/sucursales/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteSucursalesRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sucursales.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

    // XLSX
    @PostMapping("/sucursales/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteSucursalesRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sucursales.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/sucursales/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteSucursalesRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sucursales.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/sucursales/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteSucursalesRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=sucursales.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }

}
