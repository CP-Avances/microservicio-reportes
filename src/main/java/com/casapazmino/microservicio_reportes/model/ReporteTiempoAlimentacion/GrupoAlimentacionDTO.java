package com.casapazmino.microservicio_reportes.model.ReporteTiempoAlimentacion;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GrupoAlimentacionDTO {
    private String sucursal;
    private String ciudad;
    private String departamento;
    private String nombre;

    private List<EmpleadoAlimentacionDTO> empleados;

}
