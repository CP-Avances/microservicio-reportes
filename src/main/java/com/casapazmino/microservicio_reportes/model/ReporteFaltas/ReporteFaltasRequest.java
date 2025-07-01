package com.casapazmino.microservicio_reportes.model.ReporteFaltas;

import java.util.List;

public class ReporteFaltasRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    private String fechaInicio;
    private String fechaFin;
    private int opcionBusqueda;
    private boolean resumen;

    private List<GrupoFaltasDTO> grupos;

    // Getters y Setters

    public String getUsuario() {
        return usuario;
    }

    public void setUsuario(String usuario) {
        this.usuario = usuario;
    }

    public String getEmpresa() {
        return empresa;
    }

    public void setEmpresa(String empresa) {
        this.empresa = empresa;
    }

    public String getFraseMarcaAgua() {
        return fraseMarcaAgua;
    }

    public void setFraseMarcaAgua(String fraseMarcaAgua) {
        this.fraseMarcaAgua = fraseMarcaAgua;
    }

    public String getLogoBase64() {
        return logoBase64;
    }

    public void setLogoBase64(String logoBase64) {
        this.logoBase64 = logoBase64;
    }

    public String getColorPrincipal() {
        return colorPrincipal;
    }

    public void setColorPrincipal(String colorPrincipal) {
        this.colorPrincipal = colorPrincipal;
    }

    public String getColorSecundario() {
        return colorSecundario;
    }

    public void setColorSecundario(String colorSecundario) {
        this.colorSecundario = colorSecundario;
    }

    public String getFechaInicio() {
        return fechaInicio;
    }

    public void setFechaInicio(String fechaInicio) {
        this.fechaInicio = fechaInicio;
    }

    public String getFechaFin() {
        return fechaFin;
    }

    public void setFechaFin(String fechaFin) {
        this.fechaFin = fechaFin;
    }

    public int getOpcionBusqueda() {
        return opcionBusqueda;
    }

    public void setOpcionBusqueda(int opcionBusqueda) {
        this.opcionBusqueda = opcionBusqueda;
    }

    public boolean isResumen() {
        return resumen;
    }

    public void setResumen(boolean resumen) {
        this.resumen = resumen;
    }

    public List<GrupoFaltasDTO> getGrupos() {
        return grupos;
    }

    public void setGrupos(List<GrupoFaltasDTO> grupos) {
        this.grupos = grupos;
    }
}
