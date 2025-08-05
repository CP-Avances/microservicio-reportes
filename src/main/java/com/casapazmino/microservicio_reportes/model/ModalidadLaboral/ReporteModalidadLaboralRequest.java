package com.casapazmino.microservicio_reportes.model.ModalidadLaboral;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteModalidadLaboralRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<ModalidadLaboralDTO> modalidades;

    public ReporteModalidadLaboralRequest() {}

}
