package com.casapazmino.microservicio_reportes.model.Vacuna;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class VacunaDTO {

    private Integer id;
    private String nombre;

    public VacunaDTO() {
    }

    public VacunaDTO(int id, String nombre) {
        this.id = id;
        this.nombre = nombre;
    }

}
