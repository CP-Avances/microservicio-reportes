package com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado;

import java.util.List;
import java.util.Map;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteHorasExtraConsolidadoRequest {

    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;
    private String fechaInicio;
    private String fechaFin;
    private Integer opcionBusqueda;
    private Map<String, Boolean> resumen;
    private List<GrupoHorasExtraConsolidadoDT> grupos;
    private Boolean mostrarMonetizacion;

    public Boolean getMostrarMonetizacion() {
        return mostrarMonetizacion;
    }

    public void setMostrarMonetizacion(Boolean mostrarMonetizacion) {
        this.mostrarMonetizacion = mostrarMonetizacion;
    }

    public boolean isMostrarMonetizacion() {
        return Boolean.TRUE.equals(mostrarMonetizacion);
    }

}