package com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SolicitudVacacionEmpleadoDTO {

    private Long id_empleado;
    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String empleado;
    private String ciudad;
    private String sucursal;
    private String regimen;
    private String departamento;
    private String cargo;
    private String rol;

    private List<SolicitudVacacionReporteDTO> solicitudes;
}