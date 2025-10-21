package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

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
public class ReportePlanificacionRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String tipoFiltro;
    private String titulo;
    private String periodoInicio;
    private String periodoFin;
    private List<PlanificacionEmpleadoDTO> datos;
    private List<PlanificacionDetalleDTO> detalle_acciones;
    private List<Map<String, String>> nomenclatura;
}
