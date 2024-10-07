package kidneyproject.merge;

import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author kadir
 */
public class MergeFiles {
    private static String filePath;
    static ArrayList<String> columnNames = new ArrayList<String>();
    static ArrayList<String> columnType = new ArrayList<String>();
    private static File file;
    private static FileInputStream fis;
    private static XSSFWorkbook wb;
    private static XSSFSheet sheet;
    private static Iterator<Row> itr;
    private static int rowIndex;
    private static String tableStatement = "";
    private static FileWriter fileWriter;
    private static BufferedWriter outputBufferWriter;
    //private static File directoryPath = new File("C:\\work\\java\\Projectfiles\\Lisfiles_xlsx");
    private static File directoryPath = new File("C:\\Users\\Desktop\\java\\Projectfiles\\Lisfiles_xlsx");
    //List of all files and directories
    private static File filesList[];

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        filesList = directoryPath.listFiles();
        int fileCount = 0;
        for (File file : filesList) {
            System.out.println(file.getName() + " : Started");
            filePath = file.getAbsolutePath();
            processFile(filePath);
            System.out.println(columnNames);
            System.out.println(columnType);
            System.out.println(rowIndex);
            try {
                // Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/patient_data", "root", "0112358");
                Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/creatinine", "root", "0112358"); //TODO
                connection.prepareStatement("SET sql_mode" + "= ''").execute();
                creatTable(connection);
                insertData(connection);
                System.out.println((fileCount + 1) + " ------ " + file.getName() + " ------ Completed");
                Thread.sleep(500);
            } catch (SQLException | InterruptedException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static void insertData(Connection connection) {
        try {
            fileWriter = new FileWriter(filePath.replace(".xlsx", ".txt").replace("Lisfiles_xlsx", "MissingData"));
            outputBufferWriter = new BufferedWriter(fileWriter);
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        try {
            while (true) {
                Row row = sheet.getRow(rowIndex += 1);
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                String statement = "";
                //   System.out.println("--Start inserting---" + (row.getRowNum())); //todo
                while (cellIterator.hasNext()) {
                    try {
                        Cell cell = cellIterator.next();
                        String cellType = cell.getCellType().toString();
                        switch (cellType) {
                            case "NUMERIC":
                                statement += cell.getNumericCellValue() + ",";
                                break;
                            case "STRING":
                                statement += "'" + cell.getStringCellValue() + "'" + ",";
                                break;
                            case "BLANK":
                                statement += "'" + " " + "',";
                                break;
                            default:
                                // System.out.println("No type found");
                        }
                    } catch (Exception exception) {
                        System.out.println(exception.getMessage() + "\n" + "---ROW --- " + rowIndex);
                    }
                }
                statement = "insert into lis_data" + " " + "values " + "(" + statement.trim().substring(0, statement.length() - 1) + ")";
                //   System.out.println(statement); //todo
                if (statement.equals("insert into lis_data values (' ')")) {
                    break;
                }
                try {
                    connection.prepareStatement(statement).execute();
                } catch (SQLException e) {
                    System.out.println(e.getMessage() + " " + (rowIndex + 1));
                    writeToFile("Missing row: " + (1 + rowIndex)); //todo
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        try {
            outputBufferWriter.flush();
            outputBufferWriter.close();
            connection.close();
        } catch (IOException | SQLException e) {
            System.out.println(e.getMessage());
        }
    }

    private static void creatTable(Connection connection) {

        String tableSql = "CREATE TABLE IF NOT EXISTS lis_data" + " ";
        createTableStatement();
        String statement = tableSql + String.valueOf('(') + tableStatement + String.valueOf(')');
        System.out.println(statement);
        try {
            connection.createStatement().execute(statement);
        } catch (SQLException e) {
            throw new RuntimeException(e);
        }
    }

    private static void createTableStatement() {
        tableStatement = "";
        columnType.forEach(type -> {
            int index;
            switch (type) {
                case "NUMERIC":
                    index = columnType.indexOf(type);
                    tableStatement += columnNames.get(index);
                    columnType.set(index, "double,");
                    tableStatement += " " + columnType.get(index) + " ";
                    System.out.println("Name :" + columnNames.get(index) + " Type :" + type);
                    break;
                case "STRING", "BLANK":
                    index = columnType.indexOf(type);
                    tableStatement += columnNames.get(index);
                    columnType.set(index, "varchar(30) default '',");
                    tableStatement += " " + columnType.get(index) + " ";
                    break;
                default:
                    //System.out.println("No type found");
            }
        });
        tableStatement = tableStatement.substring(0, tableStatement.length() - 2);
    }

    private static void processFile(String filePath) {

        try {
            file = new File(filePath);
            fis = new FileInputStream(file);
            wb = new XSSFWorkbook(fis);
            sheet = wb.getSheetAt(0);
            itr = sheet.iterator();

            while (itr.hasNext()) {
                Row row = itr.next();
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                while (cellIterator.hasNext()) {
                    try {
                        Cell cell = cellIterator.next();
                        if (cell.getStringCellValue().toLowerCase().equals("lis result")) {
                            rowIndex = cell.getRowIndex();// ToDO
                            //  rowIndex = 55710;
                            cellIterator = row.cellIterator();
                            cellIterator.forEachRemaining(column -> columnNames.add(tableFormatString(column.getStringCellValue())));
                            row = itr.next();
                            cellIterator = row.cellIterator();
                            cellIterator.forEachRemaining(cellType -> columnType.add(cellType.getCellType().toString()));
                            return;
                        }
                    } catch (Exception exception) {
                        System.out.println(exception.getMessage());
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String tableFormatString(String column) {
        return "`" + column + "`";
    }

    private static void writeToFile(String missingLine) {

        try {
            outputBufferWriter.write(missingLine + System.lineSeparator());
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}
