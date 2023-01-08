package utilities;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelTest {

    public static void main(String[] args) throws IOException {


        String path = System.getProperty("user.dir")+"/src/test/resources/testData/Data1.xlsx";


        FileInputStream input = new FileInputStream(path);

        Workbook workbook = new XSSFWorkbook(input);

        Sheet sheet1 = workbook.getSheet("Sheet1");

        String firstName = sheet1.getRow(1).getCell(1).toString();
        System.out.println(firstName);


        System.out.println(sheet1.getRow(2).getCell(1).toString());


//        sheet1.getRow(3).getCell(0).setCellValue("Patel");
//        sheet1.getRow(3).getCell(1).setCellValue("Harsh");
//        sheet1.getRow(3).getCell(2).setCellValue("Patel@gmail.com");
//        sheet1.getRow(3).getCell(3).setCellValue("123 Main St. Chicago, IL 60656");

        sheet1.getRow(2).getCell(1).setCellValue("Kasymalievich");
        System.out.println(sheet1.getRow(2).getCell(1).toString());


        sheet1.createRow(3).createCell(0).setCellValue("Patel");
        sheet1.getRow(3).createCell(1).setCellValue("Harsh");





        FileOutputStream output = new FileOutputStream(path);


        workbook.write(output);
        output.close();




    }

}


