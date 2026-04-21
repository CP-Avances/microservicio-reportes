package com.casapazmino.microservicio_reportes.model.reporteSolicitudesPermisos;

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
public class ReporteSolicitudesPermisosRequest {

    // Branding
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

    private SeleccionSolicitudPermisoDTO seleccion;

    // Data ya consultada desde Node
    private List<SolicitudPermisoReporteDTO> solicitudes;
}