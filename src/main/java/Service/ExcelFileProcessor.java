package Service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

@Service
public class ExcelFileProcessor {

    private int ratePerHour;

    public void calculateOvertime(String inputFilePath, String outputFilePath) {

        try (FileInputStream inputStream = new FileInputStream(inputFilePath);

            Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> iterator = sheet.iterator();

            // Skip the header
            if (iterator.hasNext()) {
                iterator.next();
            }

            // Process each row
            while (iterator.hasNext()) {

                Row row = iterator.next();

                // Column B -> order
                int order = (int) row.getCell(1).getNumericCellValue();

                // Column C -> Id
                int employeeId = (int) row.getCell(2).getNumericCellValue();

                // Column D -> Name & Lastname
                String employeeName = row.getCell(3).getStringCellValue();

                // Column E F G H -> Employee's Position
                String employeePosition = row.getCell(4).getStringCellValue();

                // Column I J -> Department
                String employeeDepartment = row.getCell(5).getStringCellValue();

                // Column K -> Date
                String date = row.getCell(6).getStringCellValue();

                // Column M -> Clock's In
                String clockIn = row.getCell(7).getStringCellValue();

                // Column O -> Clock's Out
                String clockOut = row.getCell(8).getStringCellValue();

                // Column Q -> Ot before Clock's In
                String beforeClockIn = row.getCell(9).getStringCellValue();

                // Column R -> Ot after Clock's Out
                String afterClockOut = row.getCell(10).getStringCellValue();

                // Column S -> Ot for weekend
                String otWeekend = row.getCell(11).getStringCellValue();

                // Column T -> Total budget
                String budget = row.getCell(12).getStringCellValue();

                // Column U -> Note
                String note = row.getCell(13).getStringCellValue();
            }


        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
