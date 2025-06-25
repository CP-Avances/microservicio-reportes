package com.casapazmino.microservicio_reportes.model;

import java.util.List;

public class AgrupadorUsuariosDTO {

    private String sucursal;
    private String ciudad;
    private String nombre;       // Puede ser el nombre del régimen o cargo
    private String departamento;
    private List<UsuarioDTO> empleados;

    public String getSucursal() {
        return sucursal;
    }

    public void setSucursal(String sucursal) {
        this.sucursal = sucursal;
    }

    public String getCiudad() {
        return ciudad;
    }

    public void setCiudad(String ciudad) {
        this.ciudad = ciudad;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getDepartamento() {
        return departamento;
    }

    public void setDepartamento(String departamento) {
        this.departamento = departamento;
    }

    public List<UsuarioDTO> getEmpleados() {
        return empleados;
    }

    public void setEmpleados(List<UsuarioDTO> empleados) {
        this.empleados = empleados;
    }
}
