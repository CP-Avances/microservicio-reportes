package com.casapazmino.microservicio_reportes.model;

public class GeneroDTO {

    private int id;
    private String genero;

    public GeneroDTO(){
    }

    public GeneroDTO(int id, String genero){
        this.id=id;
        this.genero=genero;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getGenero() {
        return genero;
    }

    public void setGenero(String genero) {
        this.genero = genero;
    }


}
