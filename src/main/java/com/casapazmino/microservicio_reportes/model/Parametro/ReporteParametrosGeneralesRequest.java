package com.casapazmino.microservicio_reportes.model.Parametro;

import java.util.List;

public class ReporteParametrosGeneralesRequest {
    private String nombreEmpresa;
    private String nombreUsuario;
    private String fraseMarcaAgua;
    private String colorPrincipal;
    private String colorSecundario; // ← NUEVO CAMPO
    private String logoBase64;
    private List<ParametroDTO> parametros;

    public String getNombreEmpresa() {
        return nombreEmpresa;
    }

    public void setNombreEmpresa(String nombreEmpresa) {
        this.nombreEmpresa = nombreEmpresa;
    }

    public String getNombreUsuario() {
        return nombreUsuario;
    }

    public void setNombreUsuario(String nombreUsuario) {
        this.nombreUsuario = nombreUsuario;
    }

    public String getFraseMarcaAgua() {
        return fraseMarcaAgua;
    }

    public void setFraseMarcaAgua(String fraseMarcaAgua) {
        this.fraseMarcaAgua = fraseMarcaAgua;
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

    public String getLogoBase64() {
        return logoBase64;
    }

    public void setLogoBase64(String logoBase64) {
        this.logoBase64 = logoBase64;
    }

    public List<ParametroDTO> getParametros() {
        return parametros;
    }

    public void setParametros(List<ParametroDTO> parametros) {
        this.parametros = parametros;
    }
}
