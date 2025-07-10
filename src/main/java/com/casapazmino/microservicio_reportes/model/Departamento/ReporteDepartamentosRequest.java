package com.casapazmino.microservicio_reportes.model.Departamento;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteDepartamentosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<DepartamentoDTO> departamentos;

    public ReporteDepartamentosRequest() {}

}
