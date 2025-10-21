package com.casapazmino.microservicio_reportes.model.Coordenada;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class CoordenadaDTO {
    private Integer id;
    private String descripcion;
    private String latitud;
    private String longitud;
}
