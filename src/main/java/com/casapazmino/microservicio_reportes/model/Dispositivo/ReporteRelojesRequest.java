package com.casapazmino.microservicio_reportes.model.Dispositivo;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteRelojesRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<RelojDTO> relojes;

    public ReporteRelojesRequest() {
    }
}
