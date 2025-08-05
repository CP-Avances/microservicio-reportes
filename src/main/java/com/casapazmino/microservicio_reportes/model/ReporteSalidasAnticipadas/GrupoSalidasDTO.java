package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GrupoSalidasDTO {
    private String sucursal;
    private String ciudad;
    private String departamento;
    private String nombre;
    private List<EmpleadoSalidaDTO> empleados;
    
}
