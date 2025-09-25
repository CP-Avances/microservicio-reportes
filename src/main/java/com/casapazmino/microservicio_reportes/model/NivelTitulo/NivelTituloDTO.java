package com.casapazmino.microservicio_reportes.model.NivelTitulo;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class NivelTituloDTO {

    private Integer id;
    private String nombre;

    public NivelTituloDTO() {
    }

    public NivelTituloDTO(int id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

}
