import com.gembox.spreadsheet.*;

import java.time.LocalDateTime;
import java.util.Random;

public class Program {
    public static void main(String[] args) throws java.io.IOException {
        // If using Professional version, put your serial key below.
        SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");
        ExcelFile workbook = ExcelFile.load("Template.xlsx");
        int workingDays = 8;

        LocalDateTime startDate = LocalDateTime.now().plusDays(-workingDays);
        LocalDateTime endDate = LocalDateTime.now();

        ExcelWorksheet worksheet = workbook.getWorksheet(0);

        RowColumn rowColumnPosition;
        if ((rowColumnPosition = worksheet.getCells().findText("[Company Name]", true, true)) != null)
            worksheet.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue("ACME Corp");
        if ((rowColumnPosition = worksheet.getCells().findText("[Company Address]", true, true)) != null)
            worksheet.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue("240 Old Country Road, Springfield, IL");
        if ((rowColumnPosition = worksheet.getCells().findText("[Start Date]", true, true)) != null)
            worksheet.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue(startDate);
        if ((rowColumnPosition = worksheet.getCells().findText("[End Date]", true, true)) != null)
            worksheet.getCell(rowColumnPosition.getRow(), rowColumnPosition.getColumn()).setValue(endDate);

        int row = 17;
        worksheet.getRows().insertCopy(row + 1, workingDays - 1, worksheet.getRow(row));

        Random random = new Random();
        for (int i = 0; i < workingDays; i++) {
            ExcelRow currentRow = worksheet.getRow(row + i);
            currentRow.getCell(1).setValue(startDate.plusDays(i));
            currentRow.getCell(2).setValue(random.nextInt(11) + 1);
        }

       // worksheet.calculate();
       // workbook.save("Template_result.xlsx");
    }
}
