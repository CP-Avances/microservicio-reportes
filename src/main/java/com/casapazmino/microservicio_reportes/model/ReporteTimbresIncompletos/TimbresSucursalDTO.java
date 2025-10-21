package com.casapazmino.microservicio_reportes.model.ReporteTimbresIncompletos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TimbresSucursalDTO {
    private String sucursal;
    private String ciudad;
    private String nombre;
    private String departamento;
    private List<EmpleadoDTO> empleados;
}
