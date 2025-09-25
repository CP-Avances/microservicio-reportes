package com.casapazmino.microservicio_reportes.model.Genero;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class GeneroDTO {

    private Integer id;
    private String genero;

    public GeneroDTO(){
    }

    public GeneroDTO(int id, String genero){
        this.id=id;
        this.genero=genero;
    }

}
