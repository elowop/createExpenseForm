/*
 * Name: Andrew Xu 
 * Date: 3/22/2020
 * Topic: Creates expense files starting from a specific date 
 * Ver: 1.0
 */
package createexpenseform;
import java.io.File;
import java.io.IOException;
import jxl.Cell; 
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import java.util.Date; 
import java.util.Calendar;
import java.util.Scanner;
/**
 *
 * @author Andrew
 */
public class CreateExpenseForm {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, BiffException, WriteException {
        
        System.out.println("Instructions: ");
        System.out.println("When inputing the file into the program, give the entire file path");
        System.out.println("Ex. C:\\Andrew Stuff\\Dong_Expense20200404.xls");
        System.out.println("Then enter the number of copies needed to be made");
        System.out.println("------------------------------------------------------------------------");
        
        Scanner input = new Scanner(System.in);
        
        String fileName; 
        int copies; 
        fileName = findFileData();
        //find copies 
        System.out.println("Enter amount of copies needed (Ex: 5): ");
        copies = input.nextInt();   
        createExcelFile(fileName, copies);
        
    }
    
    public static void createExcelFile(String fileName, int newNumFiles) throws IOException, BiffException, WriteException
    {
        String monthString; 
        String dayString;
        String yearString;
        int index = fileName.indexOf(".");  
        String dateString = fileName.substring(index-8, index);  
        String preFileName = fileName.substring(0, index-8);        
        File file = new File(fileName); 
        Workbook wb = Workbook.getWorkbook(file);
        int year = Integer.parseInt(dateString.substring(0,4));
        int month = Integer.parseInt(dateString.substring(4,6))-1;
        int day = Integer.parseInt(dateString.substring(6));
        Calendar calendar = Calendar.getInstance();
        calendar.set(year, month, day);
        
        for(int i = 0; i < newNumFiles; i++)
        {
            calendar.add(Calendar.DATE, 7);
            yearString = Integer.toString(calendar.get(Calendar.YEAR));
            monthString = Integer.toString(calendar.get(Calendar.MONTH)+1);
            dayString = Integer.toString(calendar.get(Calendar.DATE));
            if (monthString.length() == 1)
            {
                monthString = "0".concat(monthString);
            }
            if (dayString.length() == 1)
            {
                dayString = "0".concat(dayString);
            }
            dateString = yearString + monthString + dayString;
            File file2 = new File(preFileName + dateString + ".xls");
            WritableWorkbook wbCopy = Workbook.createWorkbook(file2, wb);
            WritableSheet s = wbCopy.getSheet(0);
            setNewDate(s, yearString, monthString, dayString, calendar); 
            wbCopy.write();
            wbCopy.close();
        }
        wb.close();
    }
    
    public static void setNewDate(WritableSheet sheet, String yearString, String monthString, String dayString, Calendar currentCalendar)
    {
        String cellString; 
        cellString = sheet.getCell(1, 0).getContents().trim();
        cellString = cellString.substring(0, cellString.length()-8);
        
        cellString = cellString + yearString + monthString + dayString;

        WritableCell cell = sheet.getWritableCell(1, 0); 

        if (cell.getType() == CellType.LABEL) 
        { 
          Label l = (Label)cell; 
          l.setString(cellString); 
        }
        
        cell = sheet.getWritableCell(1, 3);
        
        if (cell.getType() == CellType.DATE)
        {
            DateTime date = (DateTime)cell;
            date.setDate(new Date(Integer.parseInt(yearString), Integer.parseInt(monthString)-1, Integer.parseInt(dayString)));
        }
       
    }
    
    public static String findFileData()
    {
        Scanner input = new Scanner(System.in);
        
        boolean fileFound = false; 
        String fileName = ""; 
        
        while(!fileFound)
        {
            try
            {
                System.out.println("Enter file name (inclue path): ");   
                fileName = input.nextLine();
                File file = new File(fileName);
                Workbook testWb = Workbook.getWorkbook(file);
                testWb.close();
                fileFound = true;
            }
            catch(Exception e)
            {
                System.out.println("File cannot be located");
            }
        }
        
        return fileName; 
    }
    
}

