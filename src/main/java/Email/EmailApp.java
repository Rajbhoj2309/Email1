package Email;


import java.io.FileReader;
import java.io.FileWriter;
import java.util.Random;
import java.util.Scanner;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class EmailApp {

    public Scanner s = new Scanner(System.in);
    // Setting up the variables
    // Defined as 'private' so that these cannot be accessed directly
    private String fname;
    private String lname;
    private String dept;
    private String email;
    private String password;
    private int mailCapacity = 10;
    private String alter_email;

    // Constructor to receive the first name and the last name
    public EmailApp(String fname, String lname) {
        this.fname = fname;
        this.lname = lname;
        System.out.println("NEW EMPLOYEE: " + this.fname + " " + this.lname);

        // Call a method asking for the department - return the department
        this.dept = this.setDept();

        // Call a method that returns a random password
        this.password = this.generate_password(8);

        // Combine elements to generate an email
        this.email = this.generate_email();
    }

//    // Generating the email according to the given syntax
    private String generate_email() {
        return this.fname.toLowerCase() + "." + this.lname.toLowerCase() + "@" + this.dept.toLowerCase()
                + ".company.com";
    }
//
//    // Ask for the department
    private String setDept() {
        System.out.println(
                "DEPARTMENT CODES\n1 for Sales\n2 for Development\n3 for Accounting\n0 for None");
        boolean flag = false;
        do {
            System.out.print("Enter Department Code: ");
            int choice = s.nextInt();
            switch (choice) {
                case 1:
                    return "Sales";
                case 2:
                    return "Development";
                case 3:
                    return "Accounting";
                case 0:
                    return "None";
                default:
                    System.out.println("**INVALID CHOICE**");
            }
        } while (!flag);
        return null;
    }

//    // Generating a random password
    private String generate_password(int length) {
        Random r = new Random();

        String Capital_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        String Small_chars = "abcdefghijklmnopqrstuvwxyz";
        String numbers = "0123456789";
        String symbols = "!@#$%&?";
        String values = Capital_chars + Small_chars + numbers + symbols;

        String password = "";
        for (int i = 0; i < length; i++) {
            password = password + values.charAt(r.nextInt(values.length()));
        }
        return password;
    }

    // Change the password
    public void set_password() {
        boolean flag = false;
        do {
            System.out.print("ARE YOU SURE YOU WANT TO CHANGE YOUR PASSWORD? (Y/N) : ");
            char choice = s.next().charAt(0);
            if (choice == 'Y' || choice == 'y') {
                flag = true;
                System.out.print("Enter current password: ");
                String temp = s.next();
                if (temp.equals(this.password)) {
                    System.out.println("Enter new password: ");
                    this.password = s.next();
                    System.out.println("PASSWORD CHANGED SUCCESSFULLY!");
                } else {
                    System.out.println("Incorrect Password!");
                }
            } else if (choice == 'N' || choice == 'n') {
                flag = true;
                System.out.println("PASSWORD CHANGE CANCELED!");
            } else {
                System.out.println("**ENTER A VALID CHOICE**");
            }
        } while (!flag);
    }
//
    // Set the mailbox capacity
    public void set_mailCap() {
        System.out.println("Current capacity = " + this.mailCapacity + "mb");
        System.out.print("Enter new capacity: ");
        this.mailCapacity = s.nextInt();
        System.out.println("MAILBOX CAPACITY CHANGED SUCCESSFULLY!");

    }
//
    // Set the alternate email
    public void alternate_email() {
        System.out.print("Enter new alternate email: ");
        this.alter_email = s.next();
        System.out.println("ALTERNATE EMAIL SET SUCCESSFULLY!");
    }

//    // Displaying the employee's information
    public void getInfo() {
        System.out.println("NAME: " + this.fname + " " + this.lname);
        System.out.println("DEPARTMENT: " + this.dept);
        System.out.println("EMAIL: " + this.email);
        System.out.println("PASSWORD: " + this.password);
        System.out.println("MAILBOX CAPACITY: " + this.mailCapacity + "mb");
        System.out.println("ALTER EMAIL: " + this.alter_email);
    }

    public void storefile() {
        try {
            FileWriter in = new FileWriter("D:\\JAVAECLIPSE\\CORE_JAVA\\Email1\\src\\Information.txt");
            in.write("First Name: "+this.fname);
            in.append("Last Name: "+this.lname);
            in.append("Email: "+this.email);
            in.append("Password: "+this.password);
            in.append("Capacity: "+this.mailCapacity);
            in.append("Alternate mail: "+this.alter_email);
            in.close();
            System.out.println("Stored..");
        }catch (Exception e){
        	System.out.println(""+ e);
        }
    }

    public void read_file() {
        try {
            FileReader f1 = new
                    FileReader("D:\\JAVAECLIPSE\\CORE_JAVA\\Email1\\src\\Information.txt");
            int i;
            while ((i = f1.read()) != -1) {
                System.out.print((char) i);
            }
            f1.close();
        } catch (Exception e) {
            System.out.println(""+e);
        }
        System.out.println();

    }
    public void storeFileExcel() {
 
    	
         
         try (FileOutputStream outputStream = new FileOutputStream("D:\\JAVAECLIPSE\\CORE_JAVA\\Email1\\src\\data.xlsx")) {
        	 XSSFWorkbook workbook = new XSSFWorkbook();
         	Sheet sheet = workbook.createSheet("PersonInfo");
         	 Row row = sheet.createRow(0);
              Cell cell1 = row.createCell(0);
              Cell cell2 = row.createCell(1);
              Cell cell3 = row.createCell(2);
              Cell cell4 = row.createCell(3);
              Cell cell5 = row.createCell(4);
              Cell cell6 = row.createCell(5);
              
              cell1.setCellValue(this.fname);
              cell2.setCellValue(this.lname);
              cell3.setCellValue(this.dept);
              cell4.setCellValue(this.email); 
              cell5.setCellValue(this.password);
              cell6.setCellValue(this.alter_email);
             workbook.write(outputStream);
             workbook.close();
             outputStream.close();
             System.out.println("Excel file created successfully!");
		} 
         
         catch (Exception e) {
        	 e.printStackTrace();
		}
        
    }
    
    public void readExcelData() {
    	 String filePath = "D:\\JAVAECLIPSE\\CORE_JAVA\\Email1\\src\\data.xlsx";
    	try (FileInputStream fileInputStream = new FileInputStream(new File(filePath))) {
            // Create workbook instance
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            // Get the first sheet of the workbook
            Sheet sheet = workbook.getSheetAt(0);
            // Iterate over rows
            for (Row row : sheet) {
                // Iterate over cells
                for (Cell cell : row) {
                    // Retrieve cell value based on cell type
                    CellType cellType = cell.getCellType();
                    if (cellType == CellType.STRING) {
                        System.out.print(cell.getStringCellValue() + "\t");
                    } else if (cellType == CellType.NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + "\t");
                    } else if (cellType == CellType.BOOLEAN) {
                        System.out.print(cell.getBooleanCellValue() + "\t");
                    } else if (cellType == CellType.BLANK) {
                        System.out.print("\t");
                    }
                }
                System.out.println();
            }
            // Close the workbook
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}