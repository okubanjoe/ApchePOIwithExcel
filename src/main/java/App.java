import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.*;

public class App {
        public static void main (String[]args) throws Exception {
            try{
            Class.forName("oracle.jdbc.driver.OracleDriver");
            Connection connect = DriverManager.getConnection(
                   "jdbc:oracle:thin:@localhost:1521:orclcdb","system","oracle"
            );

            Statement statement = connect.createStatement();
            ResultSet resultSet = statement.executeQuery("select * from poicom");
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("spread");

            XSSFRow row = spreadsheet.createRow(1);
            XSSFCell cell;
            cell = row.createCell(1);
            cell.setCellValue("FNAME");
            cell = row.createCell(2);
            cell.setCellValue("LNAME");
            cell = row.createCell(3);
            cell.setCellValue("ADDRESS");
            cell = row.createCell(4);
            cell.setCellValue("PHONE");
            int i = 2;

            while (resultSet.next()) {
                row = spreadsheet.createRow(i);
                cell = row.createCell(1);
                cell.setCellValue(resultSet.getString("fname"));
                cell = row.createCell(2);
                cell.setCellValue(resultSet.getString("lname"));
                cell = row.createCell(3);
                cell.setCellValue(resultSet.getString("address"));
                cell = row.createCell(4);
                cell.setCellValue(resultSet.getString("phone"));
                i++;
            }

            FileOutputStream out = new FileOutputStream(new File("exceldatabase.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("exceldatabase.xlsx written successfully");

            }catch(Exception e){ System.out.println(e);}

        }
    }
