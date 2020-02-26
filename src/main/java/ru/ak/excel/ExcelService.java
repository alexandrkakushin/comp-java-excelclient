package ru.ak.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import javax.jws.WebMethod;
import javax.jws.WebParam;
import javax.jws.WebService;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import lombok.Data;
import ru.ak.model.ExcelCell;
import ru.ak.model.ExcelFile;

/**
 * SOAP-сервис для работы с Excel-файлами
 * 
 */
@WebService(name = "ExcelClient", serviceName = "ExcelClient", portName = "ExcelClientPort")
public class ExcelService {

    /** Связь токена и Excel-файла */
    Map<UUID, ExcelFile> files = new HashMap<>();

    /**
     * Получения описания Excel-файла 
     * @param token токен Excel-файла (UUID)
     * @return описание Excel-файла
     */
    private ExcelFile getExcelFile(UUID token) {
        return files.get(token);
    }

    /**
     * Получение описания Excel-файла
     * @param token токен строкой
     * @return описание Excel-файла
     */
    private ExcelFile getExcelFile(String token) {
        return getExcelFile(UUID.fromString(token));
    }

    /**
     * Открытие Excel-файла (*.xls; *.xlsx)
     * @param fileName Имя файла
     * @return Response (object, isError, description)
     */
    @WebMethod
    public Response open(@WebParam(name = "fileName") String fileName) {
        Response response = new Response();
        try {
            UUID token = UUID.randomUUID();
            ExcelFile excelFile = new ExcelFile(fileName);

            response.setObject(token);
            files.put(token, excelFile);

        } catch (Exception ex) {
            response.setError(true);
            response.setDescription(ex.getLocalizedMessage());
        }

        return response;
    }

    /**
     * Закрытие Excel-файла
     * @param token токен
     * @return Response (object, isError, description)
     */
    @WebMethod
    public Response close(@WebParam(name = "token") String token) {
        Response response = new Response();

        ExcelFile excelFile = getExcelFile(token);
        if (excelFile != null) {
            try {
                excelFile.getWorkbook().close();

            } catch (IOException e) {
                response.setError(true);
                response.setDescription(e.getLocalizedMessage());

            } finally {
                files.remove(UUID.fromString(token));
            }
        } else {
            response.setError(true);
            response.setDescription(String.format("По токену %s нет связанного файла", token));
        }

        return response;
    }

    /**
     * Сохранение Excel-файла
     * @param token токен
     * @return Response (object, isError, description)
     */
    @WebMethod
    public Response save(@WebParam(name = "token") String token) {
        Response response = new Response();

        ExcelFile excelFile = getExcelFile(token);
        if (excelFile != null) {
            try {
                FileOutputStream fos = new FileOutputStream(excelFile.getFileName());
                excelFile.getWorkbook().write(fos);

            } catch (IOException ex) {
                response.setError(true);
                response.setDescription(ex.getLocalizedMessage());
            }

        } else {
            response.setError(true);
            response.setDescription(String.format("По токену %s нет связанного файла", token));
        }

        return response;
    }

    /**
     * Установка формулы
     * @param token токен
     * @param indexSheet индекс листа, нумерация начинается с 0
     * @param cellExcel "координаты" ячейки, например [R1C1], нумерация строки и столбцов с 1
     * @param formula формула
     * @return Response (object, isError, description)
     */
    @WebMethod
    public Response setFormula(@WebParam(name = "token") String token, @WebParam(name = "sheet") int indexSheet,
            @WebParam(name = "cell") ExcelCell cellExcel, @WebParam(name = "formula") String formula) {
        Response response = new Response();

        ExcelFile excelFile = getExcelFile(token);
        if (excelFile != null) {
            Workbook workbook = excelFile.getWorkbook();

            Sheet sheet = workbook.getSheetAt(indexSheet);

            int r = cellExcel.getR() - 1;
            int c = cellExcel.getC() - 1;

            Row row = sheet.getRow(r) == null ? sheet.createRow(r) : sheet.getRow(r);
            Cell cell = row.getCell(c) == null ? row.createCell(c) : row.getCell(c);

            cell.setCellFormula(formula);

            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            formulaEvaluator.evaluateAll();
        }

        return response;
    }
    
    /**
     * Ответ сервиса
     */
    @Data
    static class Response {

        /** Произвольный объект, в частности токен */
        private Object object;

        /** Результат выполнения операции */
        private boolean isError;

        /** Описание ошибки */
        private String description;

        public Response() {}

        public Response(boolean error, String description) {
            this();
            this.isError = error;
            this.description = description;
        }

        public void setObject(Object object) {
            this.object = object;
        }

        public void setError(boolean isError) {
            this.isError = isError;
        }

        public void setDescription(String description) {
            this.description = description;
        } 
    }
}