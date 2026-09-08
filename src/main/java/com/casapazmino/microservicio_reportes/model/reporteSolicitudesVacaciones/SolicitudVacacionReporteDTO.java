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

    // HISTORIAL DE APROBACIÓN
    private Long id_historial;

    private Integer orden_paso;

    private Long id_departamento_destino;

    private String departamento_aprobacion;

    private String departamento_nombre;

    private Long id_empleado_aprobador;

    private String autoriza;

    private String empleado_nombre;

    private String accion;

    private String estado_flujo;

    private String fecha_autorizacion;

    private String fecha_hora_accion;

    private String observacion;

    private String tipo_paso;

    private Boolean obligatorio;

    private String modo_aprobador;

    private String cargo_en_momento;

    private Boolean es_jefe_en_momento;

    private Boolean historial_activo;

    private List<SolicitudVacacionAprobacionDTO> aprobaciones;
}