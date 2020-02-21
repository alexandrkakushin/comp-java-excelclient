package ru.ak.model;

import lombok.Data;

@Data
public class Cell {
    
    private int r;
    private int c;

    public int getR() {
        return this.r;
    }

    public int getC() {
        return this.c;
    }
}