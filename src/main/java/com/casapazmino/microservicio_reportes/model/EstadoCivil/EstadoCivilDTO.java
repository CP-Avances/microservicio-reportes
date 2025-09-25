package com.casapazmino.microservicio_reportes.model.EstadoCivil;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EstadoCivilDTO {

    private Integer id;
    private String estadoCivil;

    public EstadoCivilDTO() {
    }

    public EstadoCivilDTO(int id, String estadoCivil) {
        this.id = id;
        this.estadoCivil = estadoCivil;
    }

}
