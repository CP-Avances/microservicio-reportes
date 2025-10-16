package com.casapazmino.microservicio_reportes.model.Discapacidad;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DiscapacidadDTO {

    private Integer id;
    private String nombre;

    public DiscapacidadDTO() {
    }

    public DiscapacidadDTO(Integer id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

}
