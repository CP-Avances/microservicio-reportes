package com.casapazmino.microservicio_reportes.model.Feriado;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteFeriadosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<FeriadoDTO> feriados;

}
