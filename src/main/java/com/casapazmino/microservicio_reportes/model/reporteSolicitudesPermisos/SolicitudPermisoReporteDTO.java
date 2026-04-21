package com.casapazmino.microservicio_reportes.model.reporteSolicitudesPermisos;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SolicitudPermisoReporteDTO {

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
    private Long codigo_solicitud;
    private Long id_tipo_permiso;
    private String tipo_permiso;

    private String fecha_solicitud;
    private String fecha_inicio;
    private String fecha_fin;

    private Double dias;
    private Double horas;
    private Integer minutos_totales;

    private Integer estado;
    private String estado_texto;
    private String autorizado;

    private Long id_empleado_aprobador;
    private String autoriza;
    private String fecha_autorizacion;
}