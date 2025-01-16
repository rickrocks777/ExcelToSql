import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class ExcelToSql {
    public List<String> convert(String tableName) {
        List<String> result = new ArrayList<>();
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            Connection conn = DriverManager.getConnection(DBConfig.host,DBConfig.username,DBConfig.password);
            Statement stmt = conn.createStatement();
            String createTable = "create table if not exists " + tableName + "(";
            File file = new File("Etl_mapping.xlsx");
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(file);
            XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
            Row header = sheet.getRow(0);
            for (int i = 0; i< header.getPhysicalNumberOfCells(); i++) {
                result.add(header.getCell(i).getStringCellValue());
                String cellValue = header.getCell(i).getStringCellValue();
                if(cellValue.contains(" ")) {
                    cellValue = cellValue.replace(" ","_");
                }
                createTable += cellValue + " varchar(30)";
                if(i== header.getPhysicalNumberOfCells()-1) {
                    createTable+=")";
                } else {
                    createTable+=",";
                }
            }
            System.out.println(createTable);
            stmt.execute(createTable);
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row currRow = sheet.getRow(i);
                String insertTable = "insert into " + tableName + " values('";
                for (int j = 0; j < currRow.getPhysicalNumberOfCells(); j++) {
                    insertTable += String.valueOf(currRow.getCell(j).getStringCellValue());
                    if(j==currRow.getPhysicalNumberOfCells()-1) {
                        insertTable += "')";
                    } else {
                        insertTable += "','";
                    }
                }
                System.out.println(insertTable);
                stmt.execute(insertTable);
            }
            stmt.close();
            conn.close();
        } catch (IOException | OpenXML4JException | SQLException | ClassNotFoundException e) {
            e.printStackTrace();
        }
        return result;
    }
}
