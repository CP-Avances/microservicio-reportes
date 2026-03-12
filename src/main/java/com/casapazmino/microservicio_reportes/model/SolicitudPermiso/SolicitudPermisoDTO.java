package com.casapazmino.microservicio_reportes.model.SolicitudPermiso;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SolicitudPermisoDTO {

    private Integer id;
    private Integer numero_permiso;

    // Fechas (ISO string)
    private String fecha_creacion;
    private String fecha_inicio;
    private String fecha_final;

    private String descripcion;

    private String tipo_permiso;
    private String tipo_descuento;

    private String hora_inicio;
    private String hora_fin;

    private Integer dias_permiso;
    private Integer minutos_totales;

    private Boolean incluir_feriados;
    private Boolean legalizado;
    private Integer estado;

    private String documento;

    // Distribución de días
    private Integer numero_dias_lunes;
    private Integer numero_dias_martes;
    private Integer numero_dias_miercoles;
    private Integer numero_dias_jueves;
    private Integer numero_dias_viernes;
    private Integer numero_dias_sabados;
    private Integer numero_dias_domingos;

    // ===============================
    // 🔹 NUEVOS CAMPOS PARA EL REPORTE
    // ===============================
    private String nombre_emple;
    private String apellido_emple;
    private String identificacion;
    private String codigo;

    private String nom_regimen;
    private String cargo;

    private String nom_empresa;
    private String nom_ciudad;
    private String nom_sucursal;
    private String nom_departamento;
}
