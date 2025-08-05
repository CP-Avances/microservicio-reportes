package com.casapazmino.microservicio_reportes.model.ReporteTimbresVirtualesMovil;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TimbreUsuarioDTO {
    private String fecha_hora_timbre_validado;
    private String fecha_hora_timbre;
    private String id_reloj;
    private String accion;
    private String observacion;
    private String longitud;
    private String latitud;
    // Getters y Setters
}
