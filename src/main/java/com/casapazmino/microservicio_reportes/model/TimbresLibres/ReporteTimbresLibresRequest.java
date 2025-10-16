package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReporteTimbresLibresRequest {
    private String usuario;           // quien genera
    private String empresa;           // EMPRESA EN MAYÚSCULAS
    private String fraseMarcaAgua;    // marca de agua
    private String logoBase64;        // logo base64 (png/jpg)
    private String colorPrincipal;    // hex "#RRGGBB" o "RRGGBB"
    private String colorSecundario;   // hex "#RRGGBB" o "RRGGBB"
    private String titulo;            // "LISTA DE TIMBRES LIBRES - ACTIVOS/INACTIVOS"
    private String tipoFiltro;        // descripción del filtro aplicado (opcional)
    private Integer opcionBusqueda;   // 1 activos / 2 inactivos, etc.
    private PeriodoDTO periodo;       // {inicio, fin}
    private List<DatoGrupoDTO> datos; // lista de grupos (sucursal/ciudad/departamento) con empleados
}
