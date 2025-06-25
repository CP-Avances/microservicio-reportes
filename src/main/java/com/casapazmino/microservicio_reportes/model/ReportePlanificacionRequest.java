package com.casapazmino.microservicio_reportes.model;

import java.util.List;
import java.util.Map;

public class ReportePlanificacionRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String tipoFiltro;
    private String titulo;
    private String periodoInicio;
    private String periodoFin;

    private List<PlanificacionEmpleadoDTO> datos;
    private List<PlanificacionDetalleDTO> detalle_acciones;
    private List<Map<String, String>> nomenclatura;

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

    public String getTipoFiltro() {
        return tipoFiltro;
    }

    public void setTipoFiltro(String tipoFiltro) {
        this.tipoFiltro = tipoFiltro;
    }

    public String getTitulo() {
        return titulo;
    }

    public void setTitulo(String titulo) {
        this.titulo = titulo;
    }

    public List<PlanificacionEmpleadoDTO> getDatos() {
        return datos;
    }

    public void setDatos(List<PlanificacionEmpleadoDTO> datos) {
        this.datos = datos;
    }

    public List<PlanificacionDetalleDTO> getDetalle_acciones() {
        return detalle_acciones;
    }

    public void setDetalle_acciones(List<PlanificacionDetalleDTO> detalle_acciones) {
        this.detalle_acciones = detalle_acciones;
    }

    public List<Map<String, String>> getNomenclatura() {
        return nomenclatura;
    }

    public void setNomenclatura(List<Map<String, String>> nomenclatura) {
        this.nomenclatura = nomenclatura;
    }

    public String getPeriodoInicio() {
        return periodoInicio;
    }

    public void setPeriodoInicio(String periodoInicio) {
        this.periodoInicio = periodoInicio;
    }

    public String getPeriodoFin() {
        return periodoFin;
    }

    public void setPeriodoFin(String periodoFin) {
        this.periodoFin = periodoFin;
    }
}
