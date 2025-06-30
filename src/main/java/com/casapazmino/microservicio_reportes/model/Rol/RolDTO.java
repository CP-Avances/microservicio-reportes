package com.casapazmino.microservicio_reportes.model.Rol;

import java.util.List;

public class RolDTO {
    private String nombre;
    private List<FuncionDTO> funciones;

    public RolDTO() {}

    public RolDTO(String nombre, List<FuncionDTO> funciones) {
        this.nombre = nombre;
        this.funciones = funciones;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public List<FuncionDTO> getFunciones() {
        return funciones;
    }

    public void setFunciones(List<FuncionDTO> funciones) {
        this.funciones = funciones;
    }
}
