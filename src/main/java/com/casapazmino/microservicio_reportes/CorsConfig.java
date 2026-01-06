package com.casapazmino.microservicio_reportes;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.*;

@Configuration
public class CorsConfig {

    @Bean
    public WebMvcConfigurer corsConfigurer() {
        return new WebMvcConfigurer() {
            @Override
            public void addCorsMappings(CorsRegistry registry) {
                registry.addMapping("/**")
                        .allowedOrigins(
                                "http://localhost:61585",
                                "http://localhost:4200",
                                "http://192.168.0.145:4200" // ejemplo: agrega el segundo puerto aquí
                )
                        .allowedMethods("GET", "POST", "PUT", "DELETE", "OPTIONS")
                        .allowedHeaders("*")
                        .exposedHeaders("Content-Disposition", "Authorization")
                        .allowCredentials(false)
                        .maxAge(3600);
            }
        };
    }
}
