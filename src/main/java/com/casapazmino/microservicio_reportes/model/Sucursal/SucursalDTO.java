package com.casapazmino.microservicio_reportes.model.Sucursal;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class SucursalDTO {

    private Long id;
    private String nombre;
    private String descripcion;

    public SucursalDTO() {}

    public SucursalDTO(Long id, String nombre, String descripcion) {
        this.id = id;
        this.nombre = nombre;
        this.descripcion = descripcion;
    }

}
