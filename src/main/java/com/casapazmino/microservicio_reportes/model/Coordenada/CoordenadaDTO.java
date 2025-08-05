package com.casapazmino.microservicio_reportes.model.Coordenada;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CoordenadaDTO {
    private int id;
    private String descripcion;
    private String latitud;
    private String longitud;

    public CoordenadaDTO() {
    }

    public CoordenadaDTO(int id, String descripcion, String latitud, String longitud) {
        this.id = id;
        this.descripcion = descripcion;
        this.latitud = latitud;
        this.longitud = longitud;
    }
}
