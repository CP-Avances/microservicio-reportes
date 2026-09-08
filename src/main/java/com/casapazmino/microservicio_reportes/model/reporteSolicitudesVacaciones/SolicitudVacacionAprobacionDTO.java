package com.casapazmino.microservicio_reportes.model.reporteSolicitudesVacaciones;

import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
public class SolicitudVacacionAprobacionDTO {

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
}