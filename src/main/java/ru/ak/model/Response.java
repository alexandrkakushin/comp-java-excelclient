package ru.ak.model;

import lombok.Data;

@Data
public class Response {

    private boolean error;
    private String description;

    public Response() {}

    public Response(boolean error, String description) {
        this();
        this.error = error;
        this.description = description;
    }

    public void setError(boolean error) {
        this.error = error;
    }

    public void setDescription(String description) {
        this.description = description;
    } 
}