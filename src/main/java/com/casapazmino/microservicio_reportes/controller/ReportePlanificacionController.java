package com.casapazmino.microservicio_reportes.controller;

import com.casapazmino.microservicio_reportes.model.PlanificacionHoraria.ReportePlanificacionRequest;
import com.casapazmino.microservicio_reportes.service.ReportePlanificacionService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;


@RestController
@RequestMapping("/api/reportes")
public class ReportePlanificacionController {

    @Autowired
    private ReportePlanificacionService reporteService;

    @PostMapping("/planificacion/pdf")
    public ResponseEntity<byte[]> generarReportePlanificacion(@RequestBody ReportePlanificacionRequest request) {
        byte[] pdf = reporteService.generarReportePDF(request);
        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=planificacion_horaria.pdf")
            .contentType(MediaType.APPLICATION_PDF)
            .body(pdf);
    }

    // === XLSX ===
    @PostMapping("/planificacion/xlsx")
    public ResponseEntity<byte[]> generarReportePlanificacionXLSX(@RequestBody ReportePlanificacionRequest request) {
        byte[] bin = reporteService.generarReporteXLSX(request);
        return ResponseEntity.ok()
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Planificacion_Horaria.xlsx")
            .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            .body(bin);
    }

    
}
