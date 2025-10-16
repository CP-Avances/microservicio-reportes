package com.casapazmino.microservicio_reportes.model.Dispositivo;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class RelojDTO {

    private String codigo;
    private String nomempresa;
    private String nomciudad;
    private String nomsucursal;
    private String nomdepar;
    private String nombre;
    private String ip;
    private int puerto;
    private String marca;
    private String modelo;
    private String serie;
    private String mac;
    private String idFabricacion;
    private String fabricante;
    private String zonaHorariaDispositivo;
    private String formatoGmtDispositivo;
    private Integer id;
    private String contrasenia;
    private String tipoConexion;
    private Integer idSucursal;
    private Integer idDepartamento;
    private String temperatura;

}
