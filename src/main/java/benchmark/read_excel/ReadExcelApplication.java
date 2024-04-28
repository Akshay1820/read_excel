package benchmark.read_excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import java.util.*;


@SpringBootApplication
public class ReadExcelApplication {

    private static final Logger log = LoggerFactory.getLogger(ReadExcelApplication.class);

    private static final Map<Integer, String> columnRow = new HashMap<>();

    public static void main(String[] args) throws IOException {
        SpringApplication.run(ReadExcelApplication.class, args);
        Sheet sheet = getSheet("Sheet1");
        setHeaderColumnRow(sheet);
        log.info(getExcelData(sheet).toString());
    }

    public static Sheet getSheet(String sheetName) throws IOException {
        Resource resource = new ClassPathResource("names.xlsx");
        File file = resource.getFile();
        try (FileInputStream fis = new FileInputStream(file)) {
            System.out.println("Reading excel file: " + file.getName());
            Workbook workbook = new XSSFWorkbook(fis);
            return workbook.getSheet(sheetName);
        } catch (IOException e) {
            e.printStackTrace();
            throw new IOException("File not found ");
        }
    }

    public static void setHeaderColumnRow(Sheet sheet) {
        Row headerRow = sheet.getRow(0);        // will assume that first row will contain column names
        Iterator<Cell> collumnNameIterator = headerRow.cellIterator();
        while (collumnNameIterator.hasNext()) {
            Cell cell = collumnNameIterator.next();
            Integer columnIndex = cell.getColumnIndex();
            String columnValue = cell.getStringCellValue();
            columnRow.put(columnIndex, columnValue);
        }
    }

    public static List<CustomerData> getExcelData(Sheet sheet) {
        List<CustomerData> customerDataList = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String firstName = getRowValue(row, "First name");
            String lastName = getRowValue(row, "Last name");
            CustomerData customerData = CustomerData.builder()
                    .firstName(firstName)
                    .lastName(lastName)
                    .build();
            customerDataList.add(customerData);
        }
        return customerDataList;
    }

    public static String getRowValue(Row row, String columnName) {
        Integer columnIndex = getColumnIndex(columnName);
        Cell cell = row.getCell(columnIndex);
        return cell.getStringCellValue();
    }

    public static Integer getColumnIndex(String columnName) {
        for (Map.Entry<Integer, String> column : columnRow.entrySet()) {
            if (columnName.equalsIgnoreCase(column.getValue())) {
                return column.getKey();
            }
        }
        return null;
    }

}
