package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class DatoGrupoDTO {

    private String sucursal;
    private String ciudad;
    private String departamento;
    private List<EmpleadoDTO> empleados;
}
