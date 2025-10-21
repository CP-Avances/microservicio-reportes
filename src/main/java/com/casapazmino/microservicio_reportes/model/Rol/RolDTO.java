package com.casapazmino.microservicio_reportes.model.Rol;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class RolDTO {
    private String nombre;
    private List<FuncionDTO> funciones;
}
