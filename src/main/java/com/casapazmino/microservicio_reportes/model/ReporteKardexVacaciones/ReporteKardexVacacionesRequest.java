package com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;
import java.util.Map;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ReporteKardexVacacionesRequest {

    // ===== Branding / datos generales =====
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private String colorSecundario;

    // ===== Contexto / filtros usados =====
    private String fechaCorte;

    // { activos: true/false, inactivos: true/false }
    private Map<String, Boolean> usuarios;
    private Map<String, Boolean> periodos;

    // 's' | 'r' | 'c' | 'd' | 'e'
    private String criterio;

    // seleccion puede traer cualquiera de estos arreglos de ids
    private SeleccionDTO seleccion;

    // ===== Contenido principal ya consultado =====
    private List<KardexEmpleadoDTO> empleados;
}