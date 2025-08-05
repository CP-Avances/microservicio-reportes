package com.casapazmino.microservicio_reportes.model.Coordenada;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteCoordenadasRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<CoordenadaDTO> coordenadas;

    public ReporteCoordenadasRequest() {
    }
}
