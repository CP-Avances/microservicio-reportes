package com.casapazmino.microservicio_reportes.model.ReporteFaltas;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class FaltaDTO {
    private String fecha;
    private Integer total_marcas;
    private Integer marcas_con_timbre;
    private Boolean justificada;
    private String estado_justificacion;
    private Boolean tiene_permiso;
    private Boolean tiene_vacacion;
    private String tipo_justificacion;
    private String nombre_justificacion;
    private String desde;
    private String hasta;
    private String detalle_justificacion;
    private String tipo_permiso;
    private String permiso_desde;
    private String permiso_hasta;
    private Long id_emple_permiso;
    private String vacacion_desde;
    private String vacacion_hasta;
    private Double numero_dias_vacacion;
    private Long id_solicitud_vacacion;
}