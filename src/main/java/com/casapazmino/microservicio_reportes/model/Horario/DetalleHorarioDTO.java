package com.casapazmino.microservicio_reportes.model.Horario;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DetalleHorarioDTO {

    private Integer orden;
    private String hora;
    private Integer tolerancia;
    private String tipoAccionShow;
    private boolean segundoDia;
    private Integer minutosAntes;
    private Integer minutosDespues;
}
