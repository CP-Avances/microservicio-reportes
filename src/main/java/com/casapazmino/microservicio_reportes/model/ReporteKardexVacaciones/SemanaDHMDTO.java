package com.casapazmino.microservicio_reportes.model.ReporteKardexVacaciones;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SemanaDHMDTO {

    private DHMDTO lunes;
    private DHMDTO martes;
    private DHMDTO miercoles;
    private DHMDTO jueves;
    private DHMDTO viernes;
    private DHMDTO sabado;
    private DHMDTO domingo;
}