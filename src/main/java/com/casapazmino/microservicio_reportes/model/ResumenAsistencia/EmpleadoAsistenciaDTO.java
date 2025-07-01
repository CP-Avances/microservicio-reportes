package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import java.util.List;

import com.fasterxml.jackson.annotation.JsonProperty;

public class EmpleadoAsistenciaDTO {

    private String identificacion;
    private String codigo;
    private String nombre;
    private String apellido;
    private String regimen;
    private String departamento;
    private String cargo;
    @JsonProperty("tLaborado")
    private List<RegistroAsistenciaDTO> tLaborado;

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

    public String getRegimen() {
        return regimen;
    }
    public void setRegimen(String regimen) {
        this.regimen = regimen;
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

    public List<RegistroAsistenciaDTO> getTLaborado() {
        return tLaborado;
    }
    public void setTLaborado(List<RegistroAsistenciaDTO> tLaborado) {
        this.tLaborado = tLaborado;
    }
}
