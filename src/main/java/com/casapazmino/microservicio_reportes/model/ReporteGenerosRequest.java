package com.casapazmino.microservicio_reportes.model;

import java.util.List;

public class ReporteGenerosRequest {

    private String usuario;         //  "Usuario Prueba"
    private String empresa;         //  "CASA PAZMIÑO"
    private String fraseMarcaAgua;  //  "FullTime"
    private String logoBase64;      // "data:image/jpeg;base64,..."
    private List<GeneroDTO> generos;

    public ReporteGenerosRequest() {
    }

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

    public List<GeneroDTO> getGeneros() {
        return generos;
    }

    public void setGeneros(List<GeneroDTO> generos) {
        this.generos = generos;
    }

}
