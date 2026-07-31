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
public class EmpleadoDTO {
    private String identificacion;
    private String nombre;
    private String apellido;
    private String codigo;
    private String correo;
    private String regimen;
    private String departamento;
    private String cargo;
    private String rol;
    private String ciudad;
    private String sucursal;
    private List<TimbreDTO> timbres;
}
