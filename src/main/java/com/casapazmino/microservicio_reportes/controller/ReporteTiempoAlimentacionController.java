package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion.ReporteTiempoAlimentacionRequest;
import com.casapazmino.microservicio_reportes.service.ReporteTiempoAlimentacionService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/reporte/tiempo-alimentacion")
public class ReporteTiempoAlimentacionController {

    @Autowired
    private ReporteTiempoAlimentacionService reporteTiempoAlimentacionService;

    @PostMapping("/pdf")
    public ResponseEntity<byte[]> generarReporteTiempoAlimentacion(@RequestBody ReporteTiempoAlimentacionRequest request) {
        byte[] pdfBytes = reporteTiempoAlimentacionService.generarReporteTiempoAlimentacionPDF(request);

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "reporte_tiempo_alimentacion.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);
    }

    // === XLSX ===
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> generarReporteTiempoAlimentacionXLSX(@RequestBody ReporteTiempoAlimentacionRequest request) {
        System.out.println("Recibida solicitud para generar reporte XLSX de Tiempo de Alimentación");
        byte[] bin = reporteTiempoAlimentacionService.generarReporteTiempoAlimentacionXLSX(request);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Tiempo_Alimentacion.xlsx")
                .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(bin);
    }
    
}
