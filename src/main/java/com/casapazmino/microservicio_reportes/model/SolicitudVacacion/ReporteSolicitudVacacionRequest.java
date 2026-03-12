package com.casapazmino.microservicio_reportes.model.SolicitudVacacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteSolicitudVacacionRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private SolicitudVacacionDTO solicitud;
    private List<AprobacionDTO> aprobaciones;
}