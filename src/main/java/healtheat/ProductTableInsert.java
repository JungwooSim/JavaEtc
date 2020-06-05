package healtheat;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;

public class ProductTableInsert {
    public static void main(String[] args) throws IOException, InvalidFormatException, SQLException {

        Connection conn = null;
        PreparedStatement pstmt = null;

        try {
            // 1. DBMS에 맞게 Driver를 로드.
            Class.forName("com.mysql.jdbc.Driver");
            //드라이버들이 읽히기만 하면 자동 객체가 생성되고 DriverManager에 등록된다.

            //2. mysql과 연결시키기
            String url = "jdbc:mysql://localhost:3306/healtheat?useSSL=false";

            conn = DriverManager.getConnection(url, "healtheat", "tlawjddn1!");
            System.out.println("Successfully Connection!");
        } catch (ClassNotFoundException e) {
            System.out.println("Failed because of not loading driver");
        } catch (SQLException e) {
            System.out.println("error : " + e);
        }
        String query = "INSERT INTO product (product_brand_id, nutrient_id, functionality_id, product_name, intake_way, shelf_life_month, manufacturing_number, functionality_text, storage_way, license_number, packing_material, intake_precaution, standard_specification, properties, shape) " +
                "values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
        // file loader
        Workbook workbook = WorkbookFactory.create(new File("./datafile/product_insert.xlsx"));

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // 3. Or you can use Java 8 forEach loop with lambda
        int i = 0;
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            System.out.println(i);

            Row row = rowIterator.next();
            if (i > 0) {
                try {
                    // 3. PreParedStatement 객체 생성, 객체 생성시 SQL 문장 저장
                    pstmt = conn.prepareStatement(query);
//
                    for (int j = 0; j < 15; j++) {
                        Cell cell = row.getCell(j);

                        if (cell != null) {
                            cell.setCellType(CellType.STRING);
                            pstmt.setString(j + 1, cell.getStringCellValue());
                        } else {
                            pstmt.setString(j + 1, "");
                        }
                    }
                    pstmt.executeUpdate();
                } catch (Exception e) {
                    System.out.println(e.getMessage());
                    System.out.println(i);
                }
            }
            i++;
        }

        // Closing the workbook
        workbook.close();
        conn.close();
    }
}
