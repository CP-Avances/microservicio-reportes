package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteVacunacionUsuariosRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String tipoFiltro;
    private String titulo;
    private List<AgrupadorVacunaUsuarioDTO> datos;

}
