package com.casapazmino.microservicio_reportes.model.ReporteSalidasAnticipadas;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class GrupoSalidasDTO {
    private String sucursal;
    private String ciudad;
    private String departamento;
    private String nombre;
    private List<EmpleadoSalidaDTO> empleados;
}
