import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelToSql {
    public List<String> convert(String tableName, String filePath) {
        List<String> result = new ArrayList<>();
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            Connection conn = DriverManager.getConnection(DBConfig.host,DBConfig.username,DBConfig.password);
            Statement stmt = conn.createStatement();
            String createTable = "create table if not exists " + tableName + "(";
            File file = new File(filePath);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(file);
            XSSFSheet sheet = xssfWorkbook.getSheetAt(0);
            Row header = sheet.getRow(0);
            for (int i = 0; i< header.getPhysicalNumberOfCells(); i++) {
                result.add(header.getCell(i).getStringCellValue());
                String cellValue = header.getCell(i).getStringCellValue();
                cellValue = fixSpecialChar(cellValue);
                if(isInteger(cellValue)) {
                    createTable += cellValue + " int";
                } else {
                    createTable += cellValue + " varchar(30)";
                }
                if(i== header.getPhysicalNumberOfCells()-1) {
                    createTable+=")";
                } else {
                    createTable+=",";
                }
            }
            System.out.println(createTable);
            stmt.execute(createTable);
            String insertStatement = "insert into " + tableName + "(";
            for (int i = 0; i < result.size(); i++) {
                insertStatement += fixSpecialChar(result.get(i));
                if(i== result.size()-1) {
                    insertStatement += ")";
                } else {
                    insertStatement += ",";
                }
            }
            insertStatement+= " values(";
            for (int i = 0; i < header.getPhysicalNumberOfCells(); i++) {
                insertStatement += "?";
                if(i==header.getPhysicalNumberOfCells()-1) {
                    insertStatement+=")";
                } else {
                    insertStatement+=",";
                }

            }
            System.out.println(insertStatement);
            PreparedStatement pstmt = conn.prepareStatement(insertStatement);
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row currRow = sheet.getRow(i);
                for (int j = 0; j < header.getPhysicalNumberOfCells(); j++) {
                    Cell currCell = currRow.getCell(j);
                    if(currCell == null) {
                        pstmt.setString(j+1,null);
                        continue;
                    }
                    CellType cellType = currCell.getCellType();
                    switch (cellType) {
                        case NUMERIC:
                            pstmt.setDouble(j+1,currCell.getNumericCellValue());
                            break;
                        case FORMULA:
                            pstmt.setObject(j+1,currCell.getCellFormula());
                            break;
                        case STRING:
                            pstmt.setString(j+1,currCell.getStringCellValue());
                            break;
                        default:
                            pstmt.setObject(j+1," ");
                    }
                }
                pstmt.execute();
            }
            stmt.close();
            conn.close();
        } catch (IOException | OpenXML4JException | SQLException | ClassNotFoundException e) {
            e.printStackTrace();
        }
        return result;
    }
    public boolean isInteger(String cellValue) {
        String[] possibleIntColumns = new String[]{"id","_id"};
        for(String possibleInt:possibleIntColumns) {
            String lower = cellValue.toLowerCase();
            if(lower.contains(possibleInt)) {
                return true;
            }
        }
        return false;
    }
    public String fixSpecialChar(String cellValue) {
        String result = cellValue;
        String[] specialChars = new String[]{"/","-","#","%","$","*","&"};
        if(cellValue.contains(" ")) {
            result = cellValue.replace(" ","_");
        }
        for(String specialChar: specialChars) {
            if(result.contains(specialChar)) {
                result = result.replace(specialChar,"");
            }
        }
        return result;
    }
}
