package com.casapazmino.microservicio_reportes.model.NivelTitulo;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteNivelesTitulosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<NivelTituloDTO> nivelesTitulos; 

    public ReporteNivelesTitulosRequest() {
    }

}
