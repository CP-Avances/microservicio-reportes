package com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class GrupoTiempoServicioDTO {

    private String sucursal;
    private String departamento;
    private String nombre;

    private List<EmpleadoTiempoServicioDTO> empleados;
}