package ru.ak.model;

import lombok.Data;

/**
 * Ячейка Excel-файла
 */
@Data
public class ExcelCell {
    
    /** Номер строки, нумерация с 1 */
    private int r;

    /** Номер колонки, нумерация с 1 */
    private int c;

    public int getR() {
        return this.r;
    }

    public int getC() {
        return this.c;
    }
}