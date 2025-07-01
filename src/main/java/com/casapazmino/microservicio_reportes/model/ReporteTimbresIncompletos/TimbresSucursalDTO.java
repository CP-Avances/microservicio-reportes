package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import lombok.Getter;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
public class TimbresSucursalDTO {
    private String sucursal;
    private String ciudad;
    private String nombre; // puede ser régimen, cargo o departamento
    private String departamento;
    private List<EmpleadoDTO> empleados;
}
