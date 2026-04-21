package com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;
import java.util.Map;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteSolicitudVacacionRequest {

    // Branding / cabecera
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    // Filtros
    private String fechaDesde;
    private String fechaHasta;

    // { activos: true/false, inactivos: true/false }
    private Map<String, Boolean> usuarios;

    // 's' | 'r' | 'd' | 'c' | 'e'
    private String criterio;

    // null | 1 | 2 | 3 | 4
    private Integer estadoSolicitud;

    // selección hecha en frontend
    private SeleccionSolicitudDTO seleccion;

    // data principal
    private List<SolicitudVacacionReporteDTO> solicitudes;
}