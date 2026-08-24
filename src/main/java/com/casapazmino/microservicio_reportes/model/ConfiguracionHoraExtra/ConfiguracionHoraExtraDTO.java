package com.casapazmino.microservicio_reportes.model.ConfiguracionHoraExtra;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ConfiguracionHoraExtraDTO {
    private Integer id;
    private String codigo;
    private String descripcion;
    private Integer tipoRegimen;
    private String nombreRegimen;
    private String tipoHoraExtra;
    private String recargoPorcentaje;
    private Boolean calcularAntesJornada;
    private Boolean calcularDespuesJornada;
    private Integer minutosMinimosTrabajados;
    private Boolean permiteCompensacionAtrasos;
    private Boolean descontarPermisosDeExtras;
    private Boolean requiereJornadaCompleta;
    private Boolean aplicaDiasLaborables;
    private Boolean aplicaFinesSemana;
    private Boolean aplicaFeriados;
    private String horaInicio;
    private String horaFinal;
    private Boolean descontarMinutosComida;
    private Boolean estado;
}