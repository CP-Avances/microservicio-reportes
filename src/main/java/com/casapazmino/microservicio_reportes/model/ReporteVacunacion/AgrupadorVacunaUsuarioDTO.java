package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class AgrupadorVacunaUsuarioDTO {

    private String sucursal;
    private String nombre;
    private String ciudad;
    private String departamento;
    private List<EmpleadoVacunaUsuarioDTO> empleados;

}
