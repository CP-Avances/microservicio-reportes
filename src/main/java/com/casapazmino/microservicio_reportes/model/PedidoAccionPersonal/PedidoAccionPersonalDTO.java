package com.casapazmino.microservicio_reportes.model.PedidoAccionPersonal;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class PedidoAccionPersonalDTO {

    private Integer id;
    private Integer idEmpleado;
    private String numeroAccionPersonal;
    private String fechaElaboracion;
    private String empleado;
    private String fechaRigeDesde;
    private String fechaRigeHasta;
    private Integer idTipoAccionPersonal;
    private String accionPersonal;
    private Integer idDetalleTipoAccion;
    private String detalleAccion;
    private String proceso;
    private Integer idVacacion;
}