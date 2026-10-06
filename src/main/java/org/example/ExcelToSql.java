package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

class ExcelToSql {
    public static void main(String[] args) throws Exception{

        String url = "jdbc:mysql://localhost:3306/student_db";
        String username = "root";
        String password = "arun@20042004";

        Connection connection = DriverManager.getConnection(url,username,password);

        FileInputStream file = new FileInputStream("details2.xlsx");
        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        String query = "insert into Students value (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        PreparedStatement statement = connection.prepareStatement(query);

        for(int i = 1; i <= sheet.getLastRowNum(); i++){
            Row row = sheet.getRow(i);

            if(row == null || row.getCell(0) == null){
                continue;
            }

            statement.setInt(1, (int) row.getCell(0).getNumericCellValue());
            statement.setString(2, row.getCell(1).getStringCellValue());
            statement.setInt(3, (int) row.getCell(2).getNumericCellValue());
            statement.setString(4, row.getCell(3).getStringCellValue());
            statement.setInt(5, (int) row.getCell(4).getNumericCellValue());
            statement.setInt(6, (int) row.getCell(5).getNumericCellValue());
            statement.setInt(7, (int) row.getCell(6).getNumericCellValue());
            statement.setInt(8, (int) row.getCell(7).getNumericCellValue());
            statement.setInt(9, (int) row.getCell(8).getNumericCellValue());
            statement.setInt(10, (int) row.getCell(9).getNumericCellValue());
            statement.setInt(11, (int) row.getCell(10).getNumericCellValue());
            statement.setInt(12, (int) row.getCell(11).getNumericCellValue());
            statement.setInt(13, (int) row.getCell(12).getNumericCellValue());
            statement.setInt(14, (int) row.getCell(13).getNumericCellValue());
            statement.setInt(15, (int) row.getCell(14).getNumericCellValue());
            statement.setInt(16, (int) row.getCell(15).getNumericCellValue());

            statement.executeUpdate();
        }

        System.out.print("Success");

        statement.close();
        connection.close();
        workbook.close();
        file.close();
    }
}