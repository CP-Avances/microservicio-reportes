package com.casapazmino.microservicio_reportes.model.Horario;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class HorarioDTO {
    private String nombre;
    private String horaTrabajo;
    private Integer minutosComida;
    private String codigo;
    private boolean noturno;
    private String documento;
    private List<DetalleHorarioDTO> detalles;
}
