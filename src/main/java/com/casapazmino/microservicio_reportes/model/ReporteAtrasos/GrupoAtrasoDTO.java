package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class GrupoAtrasoDTO {
    private String sucursal;
    private String ciudad;
    private String departamento;
    private String nombre;
    private List<EmpleadoAtrasoDTO> empleados;
}