import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.sql.*;
import java.util.Properties;
import java.util.logging.Logger;

public class App {
    private static  final Logger log = Logger.getLogger(App.class.getName());
    static Statement statement;
    static ResultSet resultSet;
    static XSSFWorkbook workbook;
    static XSSFSheet spreadsheet;
    static XSSFRow row;
    static XSSFCell cell;
    public static void main (String[]args) throws Exception {
      InputStream inputStream = App.class.getClassLoader().getResourceAsStream("config.properties");
       Properties prop = new Properties();
       prop.load(inputStream);
       String url = prop.getProperty("db.url");
       String userName = prop.getProperty("db.userid");
       String password = prop.getProperty("db.password");

        try(Connection connect = DriverManager.getConnection(url, userName, password)){

            statement = connect.createStatement();
            resultSet = statement.executeQuery("select * from poicom");
            workbook = new XSSFWorkbook();
            spreadsheet = workbook.createSheet("spread");
            row = spreadsheet.createRow(1);
            cell = row.createCell(1);
            cell.setCellValue("AYNAME");
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
            log.info("exceldatabase.xlsx written successfully");

        }catch(Exception e){ e.printStackTrace();}

    }
}
