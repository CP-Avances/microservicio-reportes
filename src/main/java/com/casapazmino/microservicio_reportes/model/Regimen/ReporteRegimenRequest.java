// ReporteRegimenRequest.java
package com.casapazmino.microservicio_reportes.model.Regimen;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteRegimenRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private List<RegimenDTO> regimen;

}
