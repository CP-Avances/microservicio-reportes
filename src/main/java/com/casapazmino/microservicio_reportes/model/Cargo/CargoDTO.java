package com.casapazmino.microservicio_reportes.model.Cargo;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties; 
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;


@Getter
@Setter
@NoArgsConstructor  // Genera un constructor sin argumentos, sin necesidad de escribirlo manualmente
@JsonIgnoreProperties(ignoreUnknown = true) // Ignorar propiedades desconocidas. De momento se puede enviar campos extras y no generara error
public class CargoDTO {
    private Long id;
    private String cargo;
}
