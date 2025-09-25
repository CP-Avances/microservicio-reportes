package com.casapazmino.microservicio_reportes.model.Ciudad;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class CiudadDTO {

    private Integer id;       
    private Integer id_prov;
    private String provincia;
    private String nombre;

    public CiudadDTO() {}

    public CiudadDTO(String provincia, String nombre) {
        this.provincia = provincia;
        this.nombre = nombre;
    }
}
