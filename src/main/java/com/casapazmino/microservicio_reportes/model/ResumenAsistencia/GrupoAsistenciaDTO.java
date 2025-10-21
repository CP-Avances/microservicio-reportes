package com.casapazmino.microservicio_reportes.model.ResumenAsistencia;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class GrupoAsistenciaDTO {

    private String sucursal;
    private String ciudad;
    private String nombre;
    private String departamento;
    private List<EmpleadoAsistenciaDTO> empleados;
}
