package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import lombok.Getter;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
public class GrupoTiempoDTO {
    private String sucursal;
    private String ciudad;
    private String departamento;
    private String nombre;
    private List<EmpleadoTiempoDTO> empleados;
}
