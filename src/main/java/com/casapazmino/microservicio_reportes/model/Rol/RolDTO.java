package com.casapazmino.microservicio_reportes.model.Rol;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RolDTO {
    private String nombre;
    private List<FuncionDTO> funciones;

    public RolDTO() {}

    public RolDTO(String nombre, List<FuncionDTO> funciones) {
        this.nombre = nombre;
        this.funciones = funciones;
    }

}
