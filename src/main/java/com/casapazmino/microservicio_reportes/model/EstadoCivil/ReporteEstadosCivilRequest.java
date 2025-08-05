package com.casapazmino.microservicio_reportes.model.EstadoCivil;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteEstadosCivilRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<EstadoCivilDTO> estadosCivil;

    public ReporteEstadosCivilRequest() {
    }
}
