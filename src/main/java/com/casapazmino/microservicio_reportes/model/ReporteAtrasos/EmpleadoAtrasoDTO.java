package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class EmpleadoAtrasoDTO {
    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String regimen;
    private String departamento;
    private String cargo;
    private String ciudad;
    private String sucursal;
    private List<AtrasoDTO> atrasos;
}