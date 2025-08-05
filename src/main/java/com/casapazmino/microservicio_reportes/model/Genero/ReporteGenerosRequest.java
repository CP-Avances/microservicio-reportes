package com.casapazmino.microservicio_reportes.model.Genero;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteGenerosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal; 
    private List<GeneroDTO> generos;

    public ReporteGenerosRequest() {
    }

}
