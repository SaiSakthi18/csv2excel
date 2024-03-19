package csv2excel;

import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class CSVtoExcelConverter {

    public static void main(String[] args) {
        if (args.length != 2) {
            System.out.println("Usage: java CSVtoExcelConverter <input_csv_file> <output_excel_file>");
            return;
        }

        String inputCsvFile = args[0];
        String outputExcelFile = args[1];

        try {
            FileInputStream inputStream = new FileInputStream(inputCsvFile);
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Sheet1");

            BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
            String line;
            int rowIdx = 0;

            while ((line = reader.readLine()) != null) {
                String[] rowData = line.split(",");
                XSSFRow row = sheet.createRow(rowIdx++);
                for (int i = 0; i < rowData.length; i++) {
                    row.createCell(i).setCellValue(rowData[i]);
                }
            }

            // Autosize the columns
            for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
                sheet.autoSizeColumn(i);
            }

            FileOutputStream outputStream = new FileOutputStream(outputExcelFile);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            reader.close();

            System.out.println("Excel file created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
