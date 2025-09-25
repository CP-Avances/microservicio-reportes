package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ModalidadLaboral.ReporteModalidadLaboralRequest;
import com.casapazmino.microservicio_reportes.service.ReporteModalidadLaboralService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteModalidadLaboralController {

    @Autowired
    private ReporteModalidadLaboralService reporteService;

    @PostMapping("/modalidad-laboral/pdf")
    public ResponseEntity<byte[]> generarReporte(@RequestBody ReporteModalidadLaboralRequest request) {
        byte[] pdfBytes = reporteService.generarReportePDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=modalidad_laboral.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

     // XLSX
    @PostMapping("/modalidad-laboral/xlsx")
    public ResponseEntity<byte[]> generarReporteXLSX(@RequestBody ReporteModalidadLaboralRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=modalidad_laboral.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/modalidad-laboral/csv")
    public ResponseEntity<byte[]> generarReporteCSV(@RequestBody ReporteModalidadLaboralRequest request) {
        byte[] bin = reporteService.generarReporteCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=modalidad_laboral.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/modalidad-laboral/xml")
    public ResponseEntity<byte[]> generarReporteXML(@RequestBody ReporteModalidadLaboralRequest request) {
        byte[] bin = reporteService.generarReporteXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=modalidad_laboral.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }

}
