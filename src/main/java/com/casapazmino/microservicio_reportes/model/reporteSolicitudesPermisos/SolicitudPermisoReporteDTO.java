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

    // ==========================
    // DATOS DEL EMPLEADO
    // ==========================
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

    // ==========================
    // DATOS DE LA SOLICITUD
    // ==========================
    private Long solicitud;
    private Long codigo_solicitud;
    private Long id_tipo_permiso;
    private String tipo_permiso;

    private String fecha_solicitud;
    private String fecha_inicio;
    private String fecha_fin;

    private String hora_inicio;
    private String timbre_inicio_permiso;
    private String hora_fin;
    private String timbre_fin_permiso;

    private Double dias;
    private Double horas;
    private Integer minutos_totales;

    private Integer estado;
    private String estado_texto;
    private String autorizado;

    // ==========================
    // CAMPOS ANTERIORES
    // Se mantienen por compatibilidad
    // ==========================
    private Long id_empleado_aprobador;
    private String autoriza;
    private String fecha_autorizacion;

    // ==========================
    // NUEVO HISTORIAL DE APROBACIÓN
    // Viene desde ea_historial_aprobacion
    // ==========================
    private Long id_historial;
    private Integer orden_paso;

    private Long id_departamento_destino;
    private String departamento_aprobacion;
    private String departamento_nombre;

    private String empleado_nombre;

    private String accion;
    private String estado_flujo;

    private String fecha_hora_accion;
    private String observacion;

    private String tipo_paso;
    private Boolean obligatorio;
    private String modo_aprobador;

    private String cargo_en_momento;
    private Boolean es_jefe_en_momento;
    private Boolean historial_activo;
}