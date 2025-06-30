package com.casapazmino.microservicio_reportes.model.PlanificacionHoraria;

import java.util.List;

public class PlanificacionEmpleadoDTO {

    private String nombre;
    private String apellido;
    private String identificacion;
    private String codigo;
    private String departamento;
    private String cargo;

    private List<PlanificacionHorarioMensualDTO> horarios;

    // Getters y Setters

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getApellido() {
        return apellido;
    }

    public void setApellido(String apellido) {
        this.apellido = apellido;
    }

    public String getIdentificacion() {
        return identificacion;
    }

    public void setIdentificacion(String identificacion) {
        this.identificacion = identificacion;
    }

    public String getCodigo() {
        return codigo;
    }

    public void setCodigo(String codigo) {
        this.codigo = codigo;
    }

    public String getDepartamento() {
        return departamento;
    }

    public void setDepartamento(String departamento) {
        this.departamento = departamento;
    }

    public String getCargo() {
        return cargo;
    }

    public void setCargo(String cargo) {
        this.cargo = cargo;
    }

    public List<PlanificacionHorarioMensualDTO> getHorarios() {
        return horarios;
    }

    public void setHorarios(List<PlanificacionHorarioMensualDTO> horarios) {
        this.horarios = horarios;
    }
}
