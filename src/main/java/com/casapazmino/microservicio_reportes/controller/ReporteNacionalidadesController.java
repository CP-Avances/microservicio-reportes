package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.Nacionalidad.ReporteNacionalidadesRequest;
import com.casapazmino.microservicio_reportes.service.ReporteNacionalidadesService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reportes")
public class ReporteNacionalidadesController {

    @Autowired
    private ReporteNacionalidadesService reporteService;

    @PostMapping("/nacionalidades/pdf")
    public ResponseEntity<byte[]> generarReporteNacionalidades(@RequestBody ReporteNacionalidadesRequest request) {
        byte[] pdfBytes = reporteService.generarReporteNacionalidadesPDF(request);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=nacionalidades.pdf")
                .contentType(MediaType.APPLICATION_PDF)
                .body(pdfBytes);
    }

      // XLSX
    @PostMapping("/nacionalidades/xlsx")
    public ResponseEntity<byte[]> generarReporteNacionalidadesXLSX(@RequestBody ReporteNacionalidadesRequest request) {
        byte[] bin = reporteService.generarReporteNacionalidadesXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=nacionalidades.xlsx")
                .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }

    // CSV
    @PostMapping("/nacionalidades/csv")
    public ResponseEntity<byte[]> generarReporteNacionalidadesCSV(@RequestBody ReporteNacionalidadesRequest request) {
        byte[] bin = reporteService.generarReporteNacionalidadesCSV(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=nacionalidades.csv")
                .contentType(MediaType.valueOf("text/csv"))
                .body(bin);
    }

    // XML
    @PostMapping("/nacionalidades/xml")
    public ResponseEntity<byte[]> generarReporteNacionalidadesXML(@RequestBody ReporteNacionalidadesRequest request) {
        byte[] bin = reporteService.generarReporteNacionalidadesXML(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=nacionalidades.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(bin);
    }


}
