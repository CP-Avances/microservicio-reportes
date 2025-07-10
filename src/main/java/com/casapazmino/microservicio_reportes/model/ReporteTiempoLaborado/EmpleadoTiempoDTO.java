package com.casapazmino.microservicio_reportes.model.ReporteTiempoLaborado;

import lombok.Getter;
import lombok.Setter;
import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

@Getter
@Setter
public class EmpleadoTiempoDTO {
    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String regimen;
    private String departamento;
    private String cargo;
    @JsonProperty("tLaborado")
    private List<RegistroTiempoDTO> tLaborado;
}
