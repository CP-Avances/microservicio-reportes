package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmpleadoAsistenciaDTO {

    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String regimen;
    private String departamento;
    private String cargo;
    @JsonProperty("tLaborado")
    private List<RegistroAsistenciaDTO> tLaborado;
 
}