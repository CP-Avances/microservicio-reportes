package com.casapazmino.microservicio_reportes.model.Vacuna;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteVacunasRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<VacunaDTO> vacunas;

    public ReporteVacunasRequest() {
    }

}
