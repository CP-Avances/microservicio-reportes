package com.casapazmino.microservicio_reportes.model.ReporteTimbresUsuarios;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GrupoTimbresDTO {
    private String sucursal;
    private String ciudad;
    private String nombre; 
    private String departamento;
    private List<EmpleadoTimbreDTO> empleados;

}
