package com.casapazmino.microservicio_reportes.model.Provincia;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ProvinciaDTO {

    private Integer id;       
    private Integer id_pais;  
    private String pais;
    private String nombre;
    

    public ProvinciaDTO() {}

    public ProvinciaDTO(String pais, String nombre) {
        this.pais = pais;
        this.nombre = nombre;
    }
}
