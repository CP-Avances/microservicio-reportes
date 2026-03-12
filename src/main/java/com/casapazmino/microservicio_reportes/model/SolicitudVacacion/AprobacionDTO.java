package com.casapazmino.microservicio_reportes.model.SolicitudVacacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class AprobacionDTO {

    private Integer orden_paso;
    private String departamento_nombre;
    private String empleado_nombre;

    private String accion;
    private String fecha_hora_accion;

    private String observacion;
    private String cargo_en_momento;
}