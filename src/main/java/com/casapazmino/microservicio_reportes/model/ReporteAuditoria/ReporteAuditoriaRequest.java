package com.casapazmino.microservicio_reportes.model.ReporteAuditoria;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteAuditoriaRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private List<AuditoriaDTO> auditorias;

}
