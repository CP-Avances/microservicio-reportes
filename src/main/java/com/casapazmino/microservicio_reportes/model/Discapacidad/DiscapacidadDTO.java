package com.casapazmino.microservicio_reportes.model.Discapacidad;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DiscapacidadDTO {

    private Long id;
    private String nombre;

    public DiscapacidadDTO() {
    }

    public DiscapacidadDTO(Long id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

}
