package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtualesMovil;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmpleadoTimbreDTO {
    private String identificacion;
    private String apellido;
    private String nombre;
    private String codigo;
    private String regimen;
    private String departamento;
    private String cargo;
    private String ciudad;
    private String sucursal;
    private List<TimbreUsuarioDTO> timbres;

}
