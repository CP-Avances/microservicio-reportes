package com.casapazmino.microservicio_reportes.model.Ciudad;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CiudadDTO {

    private String provincia;
    private String nombre;

    public CiudadDTO() {}

    public CiudadDTO(String provincia, String nombre) {
        this.provincia = provincia;
        this.nombre = nombre;
    }
}
