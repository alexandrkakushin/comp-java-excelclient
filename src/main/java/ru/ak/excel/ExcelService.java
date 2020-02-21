package ru.ak.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import javax.jws.WebMethod;
import javax.jws.WebParam;
import javax.jws.WebService;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ru.ak.model.Cell;
import ru.ak.model.Response;

@WebService(name = "ExcelClient", serviceName = "ExcelClient", portName = "ExcelClientPort")
public class ExcelService {

    Map<UUID, HSSFWorkbook> files = new HashMap<>();

    private HSSFWorkbook getWorkbook(UUID token) {
        return files.get(token);
    }

    private HSSFWorkbook getWorkbook(String token) {
        return getWorkbook(UUID.fromString(token));
    }


    @WebMethod
    public Response open(@WebParam(name = "fileName") String fileName) {
        Response response = new Response();
        try {
            UUID token = UUID.randomUUID();
            HSSFWorkbook excelWorkbook = new HSSFWorkbook(new FileInputStream(fileName));

            files.put(token, excelWorkbook);

        } catch (Exception ex) {
            response.setError(true);
            response.setDescription(ex.getLocalizedMessage());
        }

        return response;
    }

    @WebMethod
    public Response close(@WebParam(name = "token") String token) {
        Response response = new Response();

        HSSFWorkbook workbook = getWorkbook(token);
        if (workbook != null) {
            try {
                workbook.close();

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
    
    public void setFormula(String token, Cell cell, String formula) {
        HSSFWorkbook workbook = getWorkbook(token);
        if (workbook != null) {

        }
    }
}