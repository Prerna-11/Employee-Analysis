import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeAnalysis {

    public static void main(String[] args) {
        try {
            // Load the Excel file
            FileInputStream file = new FileInputStream("D:\\Projects\\Employee Analysis\\data\\Assignment_Timecard.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get the first sheet
            Iterator<Row> rowIterator = workbook.getSheetAt(0).iterator();

            // Open output.txt for writing the results
            FileOutputStream outputFile = new FileOutputStream("D:\\Projects\\Employee Analysis\\data\\output.txt");

            // Iterate through each row in the Excel sheet
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Skip header row
                if (row.getRowNum() == 0) {
                    continue;
                }

                // Assuming that the fourth column contains the "Time In" and the fifth column contains the "Time Out"
                String timeInStr = getCellValueAsString(row.getCell(2));
                String timeOutStr = getCellValueAsString(row.getCell(3));

                // Skip rows where "Time" or "Time Out" is encountered
                if (timeInStr.equalsIgnoreCase("Time") || timeOutStr.equalsIgnoreCase("Time Out")) {
                    System.out.println("Skipping row with invalid time data.");
                    continue;
                }

                // Parse time strings to Date objects
                Date timeIn = parseDate(timeInStr);
                Date timeOut = parseDate(timeOutStr);

                // Skip rows where date parsing fails
                if (timeIn == null || timeOut == null) {
                    System.out.println("Skipping row with unparseable date.");
                    continue;
                }

                // Call functions to analyze data and print to console
                int consecutiveDaysResult = consecutiveDays(row);
                int timeBetweenShiftsResult = timeBetweenShifts(timeIn, timeOut);
                int hoursInSingleShiftResult = hoursInSingleShift(timeIn, timeOut);

                // Print results to console
                System.out.println("Name: " + getCellValueAsString(row.getCell(7)));
                System.out.println("Consecutive days worked: " + consecutiveDaysResult + " days");
                System.out.println("Time between shifts: " + timeBetweenShiftsResult + " hours");
                System.out.println("Hours in single shift: " + hoursInSingleShiftResult + " hours");
                System.out.println();

                // Write results to output.txt
                String output = "Name: " + getCellValueAsString(row.getCell(7)) + "\n" +
                        "Consecutive days worked: " + consecutiveDaysResult + " days\n" +
                        "Time between shifts: " + timeBetweenShiftsResult + " hours\n" +
                        "Hours in single shift: " + hoursInSingleShiftResult + " hours\n\n";
                outputFile.write(output.getBytes());
            }

            // Close resources
            workbook.close();
            file.close();
            outputFile.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Function to check consecutive days worked
    private static int consecutiveDays(Row row) {
        // Assuming the pay cycle start and end dates are in columns 6 and 7 respectively
        Date payCycleStartDate = row.getCell(5).getDateCellValue();
        Date payCycleEndDate = row.getCell(6).getDateCellValue();

        // Calculate the number of days between pay cycle start and end dates
        long daysWorked = (payCycleEndDate.getTime() - payCycleStartDate.getTime()) / (24 * 60 * 60 * 1000);

        // Check if the employee has worked for 7 consecutive days
        return (int) daysWorked;
    }

    // Function to check time between shifts
    private static int timeBetweenShifts(Date timeIn, Date timeOut) {
        // Assuming the allowed time range is between 1 hour and 10 hours
        long timeDifference = timeOut.getTime() - timeIn.getTime();
        long hoursBetween = timeDifference / (60 * 60 * 1000);

        return (int) hoursBetween;
    }

    // Function to check hours worked in a single shift
    private static int hoursInSingleShift(Date timeIn, Date timeOut) {
        // Assuming the allowed hours range is more than 14 hours
        long timeDifference = timeOut.getTime() - timeIn.getTime();
        long hoursWorked = timeDifference / (60 * 60 * 1000);

        return (int) hoursWorked;
    }

    // Helper function to parse date strings to Date objects
    private static Date parseDate(String dateString) {
        try {
            if (!dateString.trim().isEmpty()) {
                // Check if the date is in numeric format (Excel date serial)
                if (dateString.matches("^\\d+(\\.\\d+)?$")) {
                    double numericValue = Double.parseDouble(dateString);
                    return org.apache.poi.ss.usermodel.DateUtil.getJavaDate(numericValue);
                } else {
                    return new SimpleDateFormat("dd-MM-yyyy hh:mm a").parse(dateString);
                }
            } else {
                // Handle the case where the date string is empty
                return null;
            }
        } catch (ParseException e) {
            e.printStackTrace();
            return null;
        }
    } // <-- Missing closing brace was added here

    // Helper function to get cell value as string
    private static String getCellValueAsString(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    // Handle numeric cells as needed, e.g., convert to string
                    return String.valueOf(cell.getNumericCellValue());
                default:
                    return "";
            }
        }
        return "";
    }
}