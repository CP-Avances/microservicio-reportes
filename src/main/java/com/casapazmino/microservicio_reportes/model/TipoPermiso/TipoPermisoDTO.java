package com.casapazmino.microservicio_reportes.model.TipoPermiso;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class TipoPermisoDTO {

    private Integer id;
    private String descripcion;
    private String tipoDescuento;
    private Boolean solicitaEmpleado;

    private Integer diasMaximoPermiso;
    private String horasMaximoPermiso;

    private Boolean incluirMinutosComida;
    private Integer diasAnticiparPermiso;
    private Integer crearDiasAnteriores;

    private Boolean justificar;
    private Integer diasJustificar;

    private Boolean documento;
    private Boolean legalizar;
    private Boolean contarFeriados;

    private Boolean correoCrear;
    private Boolean correoEditar;
    private Boolean correoEliminar;
    private Boolean correoPreautorizar;
    private Boolean correoAutorizar;
    private Boolean correoNegar;
    private Boolean correoLegalizar;

    private Integer sucursalId;
    private String sucursalNombre;
}