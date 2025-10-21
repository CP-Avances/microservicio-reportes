package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class AgrupadorVacunaUsuarioDTO {
    private String sucursal;
    private String nombre;
    private String ciudad;
    private String departamento;
    private List<EmpleadoVacunaUsuarioDTO> empleados;
}
