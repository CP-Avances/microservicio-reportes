package com.casapazmino.microservicio_reportes.model.Provincia;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class ProvinciaDTO {
    private Integer id;
    private Integer id_pais;  
    private String pais;
    private String nombre;
}
