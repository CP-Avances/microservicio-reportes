package com.casapazmino.microservicio_reportes.model.ReporteVacunacion;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmpleadoVacunaUsuarioDTO {

    private String identificacion;
    private String nombre;
    private String apellido;
    private String correo;
    private String genero;
    private String nacionalidad;
    private String cargo;
    private String regimen;
    private String codigo;
    private String rol;
    private String departamento;
    private String titulo;

    public String getDepartamento() {
        return departamento;
    }

    public void setDepartamento(String departamento) {
        this.departamento = departamento;
    }

    private List<VacunaUsuarioDTO> vacunas;

    public String getIdentificacion() {
        return identificacion;
    }

    public void setIdentificacion(String identificacion) {
        this.identificacion = identificacion;
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

    public String getCorreo() {
        return correo;
    }

    public void setCorreo(String correo) {
        this.correo = correo;
    }

    public String getGenero() {
        return genero;
    }

    public void setGenero(String genero) {
        this.genero = genero;
    }

    public String getNacionalidad() {
        return nacionalidad;
    }

    public void setNacionalidad(String nacionalidad) {
        this.nacionalidad = nacionalidad;
    }

    public String getCargo() {
        return cargo;
    }

    public void setCargo(String cargo) {
        this.cargo = cargo;
    }

    public String getRegimen() {
        return regimen;
    }

    public void setRegimen(String regimen) {
        this.regimen = regimen;
    }

    public String getCodigo() {
        return codigo;
    }

    public void setCodigo(String codigo) {
        this.codigo = codigo;
    }

    public String getRol() {
        return rol;
    }

    public void setRol(String rol) {
        this.rol = rol;
    }

    public List<VacunaUsuarioDTO> getVacunas() {
        return vacunas;
    }

    public void setVacunas(List<VacunaUsuarioDTO> vacunas) {
        this.vacunas = vacunas;
    }

    public String getTitulo() {
        return titulo;
    }

    public void setTitulo(String titulo) {
        this.titulo = titulo;
    }

}
