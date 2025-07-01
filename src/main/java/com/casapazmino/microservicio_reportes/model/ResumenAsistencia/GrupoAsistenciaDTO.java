package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import java.util.List;

public class GrupoAsistenciaDTO {

    private String sucursal;
    private String ciudad;
    private String nombre;
    private String departamento;
    private List<EmpleadoAsistenciaDTO> empleados;

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

    public List<EmpleadoAsistenciaDTO> getEmpleados() {
        return empleados;
    }
    public void setEmpleados(List<EmpleadoAsistenciaDTO> empleados) {
        this.empleados = empleados;
    }
}
