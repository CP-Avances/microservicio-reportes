package com.casapazmino.microservicio_reportes.model.Ciudad;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;


@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class CiudadDTO {
    private Integer id;       
    private Integer id_prov;
    private String provincia;
    private String nombre;
}
