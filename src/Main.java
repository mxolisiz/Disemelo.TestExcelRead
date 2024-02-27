import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

// Used:
// org.apache.poi:poi-ooxml:5.2.5
// org.apache.poi:poi:5.2.5

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {

    // Main method, this is the entry point of my project
    public static void main(String[] args) throws InvalidFormatException, IOException {

        // TIP Press <shortcut actionId="ShowIntentionActions"/> with your caret at the highlighted text
        // to see how IntelliJ IDEA suggests fixing it.
        // This code below change from printf to print (find out the difference)
        System.out.println("Hello and welcome!");
        String fileLocation = "C:\\Users\\CP365941\\Downloads\\Code\\Java\\Excel reader\\input data.xlsx";
        String nameOfSheet = "Sheet1";

        //Create Workbook instance holding reference to .xlsx file
        OPCPackage opcPackage = OPCPackage.open(new File(fileLocation));
        XSSFWorkbook wb = new XSSFWorkbook(opcPackage);

        // Get the excel sheet
        XSSFSheet sheet = wb.getSheet(nameOfSheet);

        // Initiate the Cell address/reference
        CellReference cellReference = new CellReference("A1");

        // Get the row for the cell
        XSSFRow row = sheet.getRow(cellReference.getCol());

        // Print out value of cell
        System.out.println(row.getCell(3));
    }




}