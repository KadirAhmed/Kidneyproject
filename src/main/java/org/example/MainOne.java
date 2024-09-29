
package org.example;

import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author kadir
 */
public class MainOne {
   private static String filePath = "C:\\Users\\marzi\\Desktop\\java\\Projectfiles\\Lisfiles_xlsx\\7018876_DM2009_Creatinine(Lab)_2010_Q2_LIS.xlsx";
    static ArrayList<String> columnNames = new ArrayList<String>();
    static ArrayList<String> columnType = new ArrayList<String>();
    private static File file;
    private static FileInputStream fis;
    private static XSSFWorkbook wb;
    private static XSSFSheet sheet;
    private static Iterator<Row> itr;
    private static int rowIndex;
    private static String tableStatement = "" ;
    private static FileWriter fileWriter;
    private static BufferedWriter outputBufferWriter;


    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        processFile(filePath);
        System.out.println(columnNames);
        System.out.println(columnType);
        try {
            Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/patient_data", "root", "0112358");
            creatTable(connection);
            insertData(connection);
        } catch (SQLException  e) {
            throw new RuntimeException(e);
        }
    }

    private static void insertData(Connection connection) {

        try {
            fileWriter = new FileWriter(filePath.replace(".xlsx",".txt").replace("Lisfiles_xlsx","MissingData"));
            outputBufferWriter = new BufferedWriter(fileWriter);
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        try {
            for(int i = 83 - 1 ; i<=  63490 - 1; i++){
               Row row = sheet.getRow(i);
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                String statement = "";
                System.out.println("--Start inserting---"+ (row.getRowNum()));
                while (cellIterator.hasNext()) {
                    try {
                        Cell cell = cellIterator.next();
                        String cellType = cell.getCellType().toString();
                        switch(cellType){
                            case "NUMERIC":
                                statement += cell.getNumericCellValue()+",";
                                break;
                            case "STRING":
                                statement += "'"+cell.getStringCellValue()+"'"+",";
                                break;
                            case "BLANK":
                                statement += "'"+" "+"',";
                                break;
                            default:
                                System.out.println("No type found");
                        }
                    }
                    catch (Exception exception){
                       System.out.println(exception.getMessage());
                    }
                }
                statement = "insert into lis_data"+" "+"values "+ "(" + statement.trim().substring(0, statement.length() - 1) +")";
                System.out.println(statement);
                try {
                    connection.prepareStatement(statement).execute();
                } catch (SQLException e) {
                    //System.out.println(e.getMessage());
                    writeToFile("Missing row: " + ( 1 + i));
                }
            }
        }
        catch(Exception e) {
            e.printStackTrace();
        }

        try {
            outputBufferWriter.flush();
            outputBufferWriter.close();
        } catch (IOException e) {
           System.out.println(e.getMessage());
        }
    }

    private static void creatTable(Connection connection) {

        String tableSql = "CREATE TABLE IF NOT EXISTS lis_data" +" ";

        createTableStatement();

        String statement = tableSql + String.valueOf('(') + tableStatement + String.valueOf(')');
        System.out.println(statement);
        try {
            connection.createStatement().execute( statement );
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

    private static void createTableStatement() {
        columnType.forEach(type->{
            int index;
            switch(type){
                case "NUMERIC":
                    index = columnType.indexOf(type);
                    tableStatement += columnNames.get(index);
                    columnType.set(index, "double,");
                    tableStatement += " " + columnType.get(index)+ " ";
                    break;
                case "STRING", "BLANK":
                    index = columnType.indexOf(type);
                    tableStatement += columnNames.get(index);
                    columnType.set(index, "varchar(30),");
                    tableStatement += " " + columnType.get(index)+ " ";
                    break;
                default:
                    System.out.println("No type found");
            }
        });
        tableStatement = tableStatement.substring(0, tableStatement.length() - 2);
    }

    private static void processFile (String filePath){

        try {
             file = new File(filePath);
             fis = new FileInputStream(file);
             wb = new XSSFWorkbook(fis);
             sheet = wb.getSheetAt(0);
             itr = sheet.iterator();

            while (itr.hasNext())
            {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                while (cellIterator.hasNext()) {
                    try {
                        Cell cell = cellIterator.next();
                        if (cell.getStringCellValue().toLowerCase().equals("lis result")) {
                            rowIndex = cell.getRowIndex();
                            cellIterator = row.cellIterator();
                            cellIterator.forEachRemaining(column -> columnNames.add(tableFormatString(column.getStringCellValue())));
                            row = itr.next();
                            cellIterator = row.cellIterator();
                            cellIterator.forEachRemaining(cellType->columnType.add(cellType.getCellType().toString()));
                            return;
                        }
                    }
                    catch (Exception exception){
                        System.out.println(exception.getMessage());
                    }
                }
            }
        }

        catch(Exception e) {
            e.printStackTrace();
        }
    }
    private static String tableFormatString(String column){
       return  "`"+column+"`";
    }

    private static void writeToFile(String missingLine){

        try {
            outputBufferWriter.write(missingLine + System.lineSeparator());
        }
        catch(Exception e) {
            System.out.println(e.getMessage());
        }
    }
}
