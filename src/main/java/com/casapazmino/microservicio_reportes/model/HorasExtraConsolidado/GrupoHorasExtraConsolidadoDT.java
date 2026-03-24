package com.casapazmino.microservicio_reportes.model.HorasExtraConsolidado;
import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class GrupoHorasExtraConsolidadoDT {
    private String sucursal;
    private String ciudad;
    private String nombre;
    private String departamento;
    private List<EmpleadoHorasExtraConsolidadoDTO> empleados;
}
