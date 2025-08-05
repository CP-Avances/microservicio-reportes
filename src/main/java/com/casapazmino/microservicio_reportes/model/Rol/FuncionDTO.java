package com.casapazmino.microservicio_reportes.model.Rol;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class FuncionDTO {

    private String pagina;
    private String accion;
    private String nombre_modulo;
    private boolean movil;

    public FuncionDTO() {
    }

}
