package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
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
