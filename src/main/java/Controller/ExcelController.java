package Controller;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@Controller
public class ExcelController {

    @PostMapping("/uploadExcel")
    public String uploadExcelFile(@RequestParam("file") MultipartFile file) {

        // Read the Excel file and process the data
        try (InputStream inputStream = file.getInputStream();
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Assuming the sheet is always the first one
            Sheet sheet = workbook.getSheetAt(0);

            // Start processing data from row number 6 (index 5)
            int rowIndex = 5;

            // Loop through each row until the row is empty
            while (sheet.getRow(rowIndex) != null) {
                Row row = sheet.getRow(rowIndex);

                // Get values from each cell (columns B to U)
                DataFormatter dataFormatter = new DataFormatter();
                String order = dataFormatter.formatCellValue(row.getCell(1));
                String employeeId = dataFormatter.formatCellValue(row.getCell(2));
                String fullName = dataFormatter.formatCellValue(row.getCell(3));
                String employeePosition = dataFormatter.formatCellValue(row.getCell(4));
                String department = dataFormatter.formatCellValue(row.getCell(5));
                // Add more columns as needed

                // TODO: Perform your calculations or data processing here

                // Move to the next row
                rowIndex++;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "redirect:/success";
    }

}
