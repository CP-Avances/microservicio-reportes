package com.casapazmino.microservicio_reportes.model.Cargo;

import java.util.List;
import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

/**
 * Modela el JSON que llega del frontend para generar el reporte de Cargos.
 */
@Getter
@Setter
@NoArgsConstructor  // Genera un constructor sin argumentos, sin necesidad de escribirlo manualmente
@JsonIgnoreProperties(ignoreUnknown = true) // Ignorar propiedades desconocidas durante la deserialización. De momento se puede enviar campos extras y no generara error
public class ReporteCargosRequest {
    private String usuario;
    private String empresa;
    private String fraseMarcaAgua;
    private String logoBase64;
    private String colorPrincipal;
    private List<CargoDTO> cargos;
}
