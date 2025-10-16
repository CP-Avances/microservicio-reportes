package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmpleadoSalidaDTO {
    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String regimen;
    private String departamento;
    private String cargo;
    private String ciudad;
    private String sucursal;
    private List<SalidaDTO> salidas;

}
