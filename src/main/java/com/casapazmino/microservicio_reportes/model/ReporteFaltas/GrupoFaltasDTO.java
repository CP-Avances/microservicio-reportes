package com.casapazmino.microservicio_reportes.model.ReporteFaltas;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GrupoFaltasDTO {

    private String sucursal;
    private String ciudad;
    private String nombre;
    private String departamento;
    private List<EmpleadoFaltasDTO> empleados;

}
