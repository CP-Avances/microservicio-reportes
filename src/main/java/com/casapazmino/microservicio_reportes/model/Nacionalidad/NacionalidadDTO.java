package com.casapazmino.microservicio_reportes.model.Nacionalidad;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class NacionalidadDTO {
    private Integer id;
    private String nombre;
}
