package org.example;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {


        try {
            readData("Tabela1a.xlsx");
        } catch (FileNotFoundException e) {
            System.out.println("Nevalidan Fajl!");
        } catch (IOException e) {
            System.out.println("Nevalidan exsel fail!");
        }



    }
    public static void readData (String relativePath) throws FileNotFoundException, IOException {



        FileInputStream inputStream = new FileInputStream(relativePath);

        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet("Test");



    }
}