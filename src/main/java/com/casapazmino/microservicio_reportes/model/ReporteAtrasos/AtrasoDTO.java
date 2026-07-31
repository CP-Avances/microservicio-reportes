package com.casapazmino.microservicio_reportes.model.ReporteAtrasos;

import com.fasterxml.jackson.annotation.JsonAlias;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@JsonIgnoreProperties(ignoreUnknown = true)
public class AtrasoDTO {

    // =========================================================
    // HORARIO Y TIMBRE
    // =========================================================
    private String fechaHorario;
    private String horaHorario;
    private String fechaTimbre;
    private String horaTimbre;

    private String fecha_hora_horario;
    private String fecha_hora_timbre;

    /**
     * Estado de la marcación:
     * R   = registro normal
     * P   = permiso
     * JHE = justificación por horas extra
     */
    private String estado_timbre;

    // =========================================================
    // JUSTIFICACIÓN GENERAL
    // Puede provenir de permiso o compensación por horas extra
    // =========================================================
    private Boolean justificado;

    /**
     * Valores esperados:
     * PERMISO
     * HORA_EXTRA
     */
    private String tipo_justificacion;

    /**
     * Texto listo para presentar:
     * PERMISO MÉDICO
     * HORAS EXTRA
     */
    private String tipo_justificacion_texto;

    /**
     * Descripción más detallada:
     * COMPENSACIÓN POR HORAS EXTRA
     * PERMISO MÉDICO
     */
    private String descripcion_justificacion;

    /**
     * Tiempo realmente aplicado a la novedad, en segundos.
     */
    private Long diferencia_justificacion;

    /**
     * Minutos/segundos compensados específicamente por horas extra.
     * Desde el frontend se envía en segundos, igual que
     * diferencia_justificacion.
     */
    private Long compensacion_hora_extra_aplicada;

    /**
     * Tiempo justificado en formato HH:mm:ss.
     */
    private String justificacion_tiempo;

    /**
     * Tiempo justificado expresado en minutos decimales.
     */
    private String justificacion_decimal;

    private Boolean es_justificacion_permiso;
    private Boolean es_justificacion_hora_extra;

    // =========================================================
    // INFORMACIÓN DEL PERMISO
    // Se conserva para compatibilidad con el formato anterior
    // =========================================================
    private String tipo_permiso;
    private String desde;
    private String hasta;

    /**
     * Duración total del permiso, en segundos.
     */
    private Long diferencia_permiso;

    /**
     * Tiempo del permiso realmente aplicado al atraso, en segundos.
     */
    private Long permiso_aplicado;

    private Long id_emple_permiso;

    @JsonAlias({ "permiso", "permiso_tiempo" })
    private String permiso_tiempo;

    private String permiso_decimal;
    private String permiso_aplicado_tiempo;
    private String permiso_aplicado_decimal;

    // =========================================================
    // ATRASO ORIGINAL
    // =========================================================
    private Long diferencia_original;
    private String atraso_original_tiempo;
    private String atraso_original_decimal;

    // =========================================================
    // ATRASO DESPUÉS DE JUSTIFICACIÓN Y ANTES DE TOLERANCIA
    // =========================================================
    private Long diferencia_antes_tolerancia;
    private String atraso_antes_tolerancia_tiempo;
    private String atraso_antes_tolerancia_decimal;

    // =========================================================
    // TOLERANCIA
    // =========================================================
    private String tolerancia;
    private Integer tolerancia_minutos;
    private String parametro_tolerancia;

    // =========================================================
    // ATRASO FINAL PENDIENTE
    // =========================================================
    private Long diferencia;
    private String tiempoAtraso;
    private String minutosAtraso;
}