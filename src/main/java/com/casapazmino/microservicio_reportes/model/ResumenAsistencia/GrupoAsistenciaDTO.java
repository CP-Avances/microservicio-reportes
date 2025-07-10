package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GrupoAsistenciaDTO {

    private String sucursal;
    private String ciudad;
    private String nombre;
    private String departamento;
    private List<EmpleadoAsistenciaDTO> empleados;
 
}