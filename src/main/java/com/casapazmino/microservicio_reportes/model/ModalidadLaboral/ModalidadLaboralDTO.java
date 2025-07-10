package com.casapazmino.microservicio_reportes.model.ModalidadLaboral;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ModalidadLaboralDTO {

    private Long id;
    private String descripcion;

    public ModalidadLaboralDTO() {}

    public ModalidadLaboralDTO(Long id, String descripcion) {
        this.id = id;
        this.descripcion = descripcion;
    }

}
