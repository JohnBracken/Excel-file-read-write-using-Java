/*The folling code shows an example of how to read from and write data to a 
Microsoft Excel spreadsheet file.  This is could be useful for applications
that require manipulation,calculation and evaluation of data stored elsewhere.
For this code to work, it required the installation of the Apache POI
libraries and JAR files.*/
package excelreadwrite;

//Libraries needed for writing to and reading from other files.
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;


/*With the Apache POI 3.15 libary loaded, these classes are needed to create
and work with Excel spreadsheets.*/
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



//Class definition for reading and writing to Excel.
public class ExcelReadWrite 
{
    
    /*Main method.  In this case, since file input and output streams are
    being used, the option to be able to throw an input/output exception has
    been provided if something goes wrong with streaming data.*/
    public static void main(String[] args) throws IOException  
    {
        //Create the spreadsheet workbook object.
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        //Create a new worksheet in the spreadsheet workbook.
        XSSFSheet sheet = workbook.createSheet("Rock Albums");
        
        /*This code creates an object of table data for the spreadsheet.  It
        is basically a 2D array of data that will go into the spreadsheet.
        In this case, it is a list of rock albums along with their year of
        release.*/ 
        Object[][] ExcelData = 
        {
                {"Band     ", "Album          ", "        Year Released"},
                {" " ," " , " "},
                {"Helix    ", "Walkin' the Razors Edge", 1984},
                {"Nine Inch Nails", "Hesitation Marks", 2013},
                {"Weezer    ", "The Blue Album    ", 1994},
                {"Metallica", "Master of Puppets", 1986},
                {"Soundgarden", "Badmotorfinger    ", 1991},        
        };
         
        
        
         //Start creating rows in the spreadsheet with a for loop.
         for (int row_pos=0; row_pos< ExcelData.length; row_pos++) {
            Row row = sheet.createRow(row_pos);
             
           
            /*In each row of the spreadsheet data, create a cell for each
            column with the inner for loop through the Excel data object.*/
            for (int col_pos=0; col_pos <ExcelData[row_pos].length; col_pos++) 
            {
                //Create a cell at the current row, column position.
                Cell cell = row.createCell(col_pos);
                
                /*Look at the Excel data object at the current (row, column)
                position.  If the data in that position is a string, then
                put that string into the spreadsheet object. */  
                if (ExcelData[row_pos][col_pos] instanceof String) 
                {
                    cell.setCellValue((String)ExcelData[row_pos][col_pos]);
                } 
                
                /*If the data in the current position of the Excel data object
                 is an integer instead, then put the data into the spreadsheet
                 object at that cell position as an integer.*/ 
                else if (ExcelData[row_pos][col_pos] instanceof Integer)     
                {
                    cell.setCellValue((Integer) ExcelData[row_pos][col_pos]);
                }
            }
             
        }
         
        /*Create a new file output stream, which will be the Excel file
        that the data is delivered to.*/
         try(FileOutputStream outputStream = 
                 new FileOutputStream("Rock_Albums.xlsx"))
         {
          
          //Write the spreadsheet object to the Excel file. 
          workbook.write(outputStream);
          
         }
         
         
         
         /*Now this section of code will read in the data that is now
         present in the newly created spreadsheet. */
         try
         
         //Create a new input stream to read in the data from the Excel file.    
         (FileInputStream inputStream = new FileInputStream("Rock_Albums.xlsx"))
         {
 
            /*Create the workbook to hold the data read in from the Excel
            file.*/
            XSSFWorkbook workbook_read = new XSSFWorkbook(inputStream);
 
            //Setup the workbook sheet to read in the data.
            XSSFSheet sheet_read = workbook_read.getSheetAt(0);
 
            //Iterate through each of the rows in the worksheet
            for (int row_pos=0; row_pos < (sheet_read.getLastRowNum() + 1); row_pos++) 
            {

                //Create a row to read the data from the excell file
                Row row_read =  sheet_read.getRow(row_pos);
                
                //Iterate through the columns in each row of the worksheet.  
                for (int col_pos=0; col_pos< row_read.getLastCellNum(); col_pos++)
                {
                    //Get the contents of the current cell in the worksheet.
                    Cell cell_read = row_read.getCell(col_pos);
                    
                    /*Look at the value that has just been read into the
                    input worksheet cell and check the type of data that 
                    is there.  The getCellType() and getCellTypeEnum() methods
                    are normally used to do this but they have currently been 
                    deprecated in Apache POI 3.15.  Alternative methods may
                    be developed in future releases of Apache POI.  For now, 
                    if this method returns a value of 1, then the cell in the
                    spreadsheet is either a string or blank.  If it has a value
                    of zero, then the cell contains a numeric value.  If the
                    method returns one, then get the value of the string in the
                    cell and print it out.*/  
                    if (cell_read.getCellType() == 1)
                    {
                    System.out.print(cell_read.getStringCellValue() + "\t\t");
                    } 
                    
                    /*If the data in the current position of the Excel data object
                    is an number instead, then print the worksheet cell
                    as an number.  Since we are dealing with years in this 
                    case, cast the number as an integer before printing it.*/ 
                    else if (cell_read.getCellType() == 0)     
                    {
                    System.out.print((int)cell_read.getNumericCellValue() + "\t\t");
                    }
                }
                //Print out a blank space to move down to the next row.
                System.out.println(" ");
            }
            
            /*Close the file stream when done reading the data from the Excel
            file. */
            inputStream.close();
        
            }
    }
}
    

