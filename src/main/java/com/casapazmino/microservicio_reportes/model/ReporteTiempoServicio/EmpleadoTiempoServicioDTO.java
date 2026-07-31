package com.casapazmino.microservicio_reportes.model.ReporteTiempoServicio;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class EmpleadoTiempoServicioDTO {

    private Long id;

    private String identificacion;
    private String codigo;

    private String nombre;
    private String apellido;

    /*
     * Datos organizacionales generales del empleado.
     * Funcionan como respaldo si algún contrato no tiene
     * el texto correspondiente.
     */
    private String sucursal;
    private String departamento;
    private String cargo;
    private String regimen;

    @JsonProperty("tiempoServicio")
    private List<ContratoTiempoServicioDTO> tiempoServicio;
}