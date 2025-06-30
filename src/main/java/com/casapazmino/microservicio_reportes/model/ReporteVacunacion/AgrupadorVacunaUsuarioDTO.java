package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;

public class AgrupadorVacunaUsuarioDTO {

    private String sucursal;
    private String nombre;
    private String ciudad;
    private String departamento;
    private List<EmpleadoVacunaUsuarioDTO> empleados;

    public String getSucursal() {
        return sucursal;
    }

    public void setSucursal(String sucursal) {
        this.sucursal = sucursal;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getCiudad() {
        return ciudad;
    }

    public void setCiudad(String ciudad) {
        this.ciudad = ciudad;
    }

    public String getDepartamento() {
        return departamento;
    }

    public void setDepartamento(String departamento) {
        this.departamento = departamento;
    }

    public List<EmpleadoVacunaUsuarioDTO> getEmpleados() {
        return empleados;
    }

    public void setEmpleados(List<EmpleadoVacunaUsuarioDTO> empleados) {
        this.empleados = empleados;
    }
}
