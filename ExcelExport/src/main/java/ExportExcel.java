import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExportExcel {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Person");
        Object[][] PersonData = {
                {"Saleh", "Mirzeliyev", 79},
                {"Saleh1", "Mirzeliyev1", 36},
                {"Saleh2", "Mirzeliyev2", 42},
                {"Saleh3", "Mirzeliyev3", 35},
        };
        int rowCount = 0;
        for (Object[] person : PersonData) {
            Row row = sheet.createRow(++rowCount);
            int columnCount = 0;
            for (Object field : person) {
                Cell cell = row.createCell(++columnCount);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
            try (FileOutputStream outputStream = new FileOutputStream("Sample.xlsx")) {
                workbook.write(outputStream);
            } catch (Exception e){
                e.printStackTrace();
            }
        }
    }
}
