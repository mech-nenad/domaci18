package org.example;

import com.github.javafaker.Faker;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Main {
    public static void main(String[] args) {


        try {
            readData("Tabela1a.xlsx");
        } catch (FileNotFoundException e) {
            System.out.println("Invalid path!");
        } catch (IOException e) {
            System.out.println("Invalid excel file!");
        }

        try {
            writeDataInWorkbook("Tabela1a.xlsx");
        } catch (FileNotFoundException e) {
            System.out.println("Invalid path!");
        }catch (IOException e) {
            System.out.println("Invalid excel file!");
        }



    }
    public static void readData (String relativePath) throws FileNotFoundException, IOException {



        FileInputStream inputStream = new FileInputStream(relativePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet("test");

        for (int i = 0; i < 5; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cellName = row.getCell(0);
            XSSFCell callSurname = row.getCell(1);
            String name = cellName.getStringCellValue();
            String surname = callSurname.getStringCellValue();

            System.out.println(name + " " + surname);
        }
    }

    public static void writeDataInWorkbook (String relativePath) throws IOException {
        Faker faker = new Faker();
        FileInputStream inputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook1.getSheet("test");


        for (int i = 5; i < 10; i++) {
            XSSFRow row = sheet.createRow(i);
            XSSFCell cellName = row.createCell(0);
            cellName.setCellValue(faker.name().firstName());
            XSSFCell cellSurname = row.createCell(1);
            cellSurname.setCellValue(faker.name().lastName());


            System.out.println(cellName + " " + cellSurname);

        }

        FileOutputStream outputStream = new FileOutputStream("Tabela1a.xlsx");
        workbook1.write(outputStream);
        outputStream.close();
    }

}



