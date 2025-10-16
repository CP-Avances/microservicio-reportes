package com.casapazmino.microservicio_reportes.model.TimbresLibres;

import java.util.List;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class DatoGrupoDTO {
    private String sucursal;      // puede ser null
    private String ciudad;        // puede ser null
    private String departamento;  // puede ser null
    private List<EmpleadoDTO> empleados;
}