package com.casapazmino.microservicio_reportes.model.Discapacidad;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DiscapacidadDTO {
    private Integer id;
    private String nombre;
}