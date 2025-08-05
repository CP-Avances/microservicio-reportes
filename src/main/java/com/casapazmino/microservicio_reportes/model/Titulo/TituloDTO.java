package com.casapazmino.microservicio_reportes.model.Titulo;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class TituloDTO {

    private int id;
    private String nivel;
    private String nombre;

    public TituloDTO() {
    }

    public TituloDTO(int id, String nivel, String nombre) {
        this.id = id;
        this.nivel = nivel;
        this.nombre = nombre;
    }

}
