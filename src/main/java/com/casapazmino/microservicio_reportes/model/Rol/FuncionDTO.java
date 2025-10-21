package com.casapazmino.microservicio_reportes.model.Rol;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class FuncionDTO {

    private String pagina;
    private String accion;
    private String nombre_modulo;
    private Boolean movil;
}
