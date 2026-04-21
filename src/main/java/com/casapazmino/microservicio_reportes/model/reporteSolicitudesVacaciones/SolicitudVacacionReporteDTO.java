package com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SolicitudVacacionReporteDTO {

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

    private Long solicitud;
    private Long id_periodo_vacacion;

    private String desde;
    private String hasta;
    private String fecha_registro;
    private String fecha_actualizacion;

    private Integer dias_l_v;
    private Integer dias_s_d;
    private Double dias_totales;
    private Integer minutos_totales;

    private Integer estado;
    private String estado_texto;
    private String autorizado;

    private Long id_empleado_aprobador;
    private String autoriza;
    private String fecha_autorizacion;
}