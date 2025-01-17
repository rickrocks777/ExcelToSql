import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.TreeMap;

public class SqlToExcel {
    public String convert(String tableName) {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            Connection conn = DriverManager.getConnection(DBConfig.host,DBConfig.username,DBConfig.password);
            Statement stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery("select * from " + tableName);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
            XSSFSheet sheet = xssfWorkbook.createSheet("sheet");
            TreeMap<Integer, List<Object>> map = new TreeMap<>();
            int count = 0;
            while (rs.next()) {
                List<Object> list = new ArrayList<>();
                for (int i = 1; i <= rs.getMetaData().getColumnCount(); i++) {
                    list.add(rs.getObject(rs.getMetaData().getColumnName(i)));
                }
                map.put(count,list);
                count++;
            }
            Set<Integer> keyset = map.keySet();
            int rownum =1;
            for (Integer key: keyset) {
                Row row = sheet.createRow(rownum++);
                List<Object> vals = map.get(key);
                int cellNum = 0;
                for (Object val: vals) {
                    Cell cell = row.createCell(cellNum++);
                    if(val instanceof Integer) {
                        cell.setCellValue((Integer) val);
                    }
                    if(val instanceof String) {
                        cell.setCellValue((String) val);
                    }
                }

            }
            FileOutputStream fos = new FileOutputStream("out.xls");
            xssfWorkbook.write(fos);
            fos.close();
            System.out.println(map);
        } catch (ClassNotFoundException | SQLException | IOException e) {
            e.printStackTrace();
        }
        return "success";
    }
}
