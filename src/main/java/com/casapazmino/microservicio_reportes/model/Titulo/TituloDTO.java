package com.casapazmino.microservicio_reportes.model.Titulo;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TituloDTO {
    private Integer id;
    private String nivel;
    private String nombre;
}
