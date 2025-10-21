package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TimbreDTO {

    private String fechaServidor;
    private String horaServidor;
    private String fechaDispositivo;
    private String horaDispositivo;
    private String id_reloj;
    private String accion;
    private String observacion;
    private String latitud;
    private String longitud;
}
