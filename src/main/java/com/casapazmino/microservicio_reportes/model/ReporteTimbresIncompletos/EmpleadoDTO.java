package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import lombok.Getter;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
public class EmpleadoDTO {
    private String identificacion;
    private String nombre;
    private String apellido;
    private String codigo;
    private String regimen;
    private String departamento;
    private String cargo;
    private List<TimbreDTO> timbres;
}

