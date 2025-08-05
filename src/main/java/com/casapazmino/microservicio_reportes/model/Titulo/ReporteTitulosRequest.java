package com.casapazmino.microservicio_reportes.model.Titulo;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteTitulosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<TituloDTO> titulos;

    public ReporteTitulosRequest() {
    }

}
