package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class PeriodoDTO {
    private String inicio;  // yyyy-mm-dd
    private String fin;     // yyyy-mm-dd
}