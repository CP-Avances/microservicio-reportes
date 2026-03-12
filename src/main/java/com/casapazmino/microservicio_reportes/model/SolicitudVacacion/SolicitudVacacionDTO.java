package com.casapazmino.microservicio_reportes.model.SolicitudVacacion;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class SolicitudVacacionDTO {

    private Integer id;
    private Integer estado;

    // Fechas (ISO string)
    private String fecha_inicio;
    private String fecha_final;

    // Estas pueden venir o no (según tu payload)
    private String fecha_registro;
    private String fecha_actualizacion;

    private Boolean incluir_feriados;
    private String documento;

    // Distribución de días
    private Integer numero_dias_lunes;
    private Integer numero_dias_martes;
    private Integer numero_dias_miercoles;
    private Integer numero_dias_jueves;
    private Integer numero_dias_viernes;
    private Integer numero_dias_sabado;   // tu frontend usa singular
    private Integer numero_dias_domingo;  // tu frontend usa singular
    private Integer numero_dias_totales;

    // ===============================
    // 🔹 CAMPOS “IGUAL QUE PERMISO”
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