package com.casapazmino.microservicio_reportes.model.Cargo;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CargoDTO {

    private Long id;
    private String cargo;

    public CargoDTO() {}

    public CargoDTO(Long id, String cargo) {
        this.id = id;
        this.cargo = cargo;
    }
}
