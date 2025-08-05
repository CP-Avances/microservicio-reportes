package com.casapazmino.microservicio_reportes.model.Nacionalidad;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteNacionalidadesRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<NacionalidadDTO> nacionalidades;

}
