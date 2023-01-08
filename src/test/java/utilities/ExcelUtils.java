package utilities;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtils {


    private  static Sheet sheet;
    private  static Workbook workbook;
    private  static FileInputStream inputStream;
    private  static FileOutputStream outputStream;
    private  static String path;


    public  static  void openExcelFile(String fileName, String sheetName)  {


      path = System.getProperty("user.dir")+"/src/test/resources/testData/"+fileName+".xlsx";



      try{

          inputStream = new FileInputStream(path);
          workbook = new XSSFWorkbook(inputStream);
          sheet = workbook.getSheet(sheetName);

      }catch (FileNotFoundException e){

          System.out.println("Excel spread sheet PATH is invalid: "+path);

      }catch (IOException e){
          System.out.println("Could not open Excel");
      }




    }



    public static  String getValue(int rowNum, int cellNum){

        return  sheet.getRow(rowNum).getCell(cellNum).toString();

    }


    public static void setValue(int rowNum,int  cellNum, String value)  {

        if(sheet.getPhysicalNumberOfRows()<=rowNum){
            sheet.createRow(rowNum).createCell(cellNum).setCellValue(value);
        }else if(sheet.getRow(rowNum).getPhysicalNumberOfCells()<=cellNum){
            sheet.getRow(rowNum).createCell(cellNum).setCellValue(value);
        }else {
            sheet.getRow(rowNum).getCell(cellNum).setCellValue(value);
        }



try{
    outputStream = new FileOutputStream(path);
    workbook.write(outputStream);
}catch (FileNotFoundException e){
    System.out.println("Excel spreadsheet path is invalid "+path);
}catch (IOException e) {
    System.out.println("Could not save changes to Excel");
}finally {
    try {
        outputStream.close();
    }catch (IOException e){
        System.out.println("Could not close fileOutputStream");
    }
}



    }



    public static void main(String[] args) throws IOException {
        openExcelFile("Data1","Sheet1");
        setValue(4, 0, "Hello");
        setValue(4, 1, "World");


        System.out.println(sheet.getRow(4).getCell(0).toString());
        System.out.println(sheet.getRow(4).getCell(1).toString());


    }



}


