package com.casapazmino.microservicio_reportes.model.Cargo;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteCargosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<CargoDTO> cargos;

    public ReporteCargosRequest() {}

}
