package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class AtrasoDTO {
    private String fechaHorario;
    private String horaHorario;
    private String fechaTimbre;
    private String horaTimbre;

    private String tipo_permiso;
    private String desde;
    private String hasta;
    private String permiso;
    private String descripcion_permiso;

    private String tolerancia;
    private String tiempoAtraso;
    private String minutosAtraso;
}
