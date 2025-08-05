package com.casapazmino.microservicio_reportes.model.RegimenLaboral;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class ReporteRegimenesRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private List<RegimenDTO> regimenes;
}