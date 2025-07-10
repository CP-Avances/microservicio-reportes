package com.casapazmino.microservicio_reportes.model.Provincia;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteProvinciasRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<ProvinciaDTO> provincias;

    public ReporteProvinciasRequest() {}

}
