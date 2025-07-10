package com.casapazmino.microservicio_reportes.model.Nacionalidad;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class NacionalidadDTO {

    private int id;
    private String nombre;

    public NacionalidadDTO() {
    }

    public NacionalidadDTO(int id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }
    
}
