package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtuales;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TimbreUsuarioDTO {
    private String fecha_hora_timbre_validado;
    private String fecha_hora_timbre;
    private String id_reloj;
    private String accion;
    private String observacion;
    private String longitud;
    private String latitud;
}
