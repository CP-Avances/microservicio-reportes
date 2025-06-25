package com.casapazmino.microservicio_reportes.model;

import java.util.List;

public class ReporteCiudadesRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<CiudadDTO> ciudades;

    public ReporteCiudadesRequest() {}

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

    public List<CiudadDTO> getCiudades() {
        return ciudades;
    }

    public void setCiudades(List<CiudadDTO> ciudades) {
        this.ciudades = ciudades;
    }
}
