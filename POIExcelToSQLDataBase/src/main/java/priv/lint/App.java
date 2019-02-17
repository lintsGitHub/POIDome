package priv.lint;

import com.mchange.v2.c3p0.ComboPooledDataSource;
import com.mchange.v2.c3p0.impl.C3P0PooledConnection;
import com.sun.xml.internal.messaging.saaj.packaging.mime.util.LineInputStream;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.beans.PropertyVetoException;
import java.io.*;
import java.sql.Connection;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 */
public class App {
    public static ComboPooledDataSource comboPooledDataSource;

    public static void main(String[] args) {
        String path = "D:/1.xls";
        try {
            List<Enterprise> enterprises = allEnterprise(path);
//            System.out.println(enterprises.size());
            Statement statement = writeDataBase();
            for (Enterprise enterpris : enterprises) {
                String sql = "insert into ge_enterprise_category values (" + enterpris.getId() + ",'" + enterpris.getName() + "'," + enterpris.getParentId() + ")";
                statement.addBatch(sql);
            }
            statement.executeBatch();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (PropertyVetoException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }

    }

    //    计数
    static int count;

    //    使用POI到Excel中拿数据
    public static List<Enterprise> allEnterprise(String path) throws IOException {
        File excelFile = null;  //Excel文件对象
        InputStream inputStream = null; //文件输入流对象
        List<Enterprise> enterprises = new ArrayList<>(); //企业信息集合
        excelFile = new File(path);
        inputStream = new FileInputStream(excelFile);
        HSSFWorkbook sheets = new HSSFWorkbook(inputStream);
        HSSFSheet sheet = sheets.getSheetAt(0);
        for (int i = 3; i < 164; i++) {
            HSSFRow row = sheet.getRow(i);

            Enterprise enterprise = new Enterprise();
            int id = (int) row.getCell(0).getNumericCellValue();
            enterprise.setId(id);
            enterprise.setName(row.getCell(1).toString().trim());
            if (id % 1000 == 0)
                count = id;
            else
                enterprise.setParentId(id);
            enterprises.add(enterprise);
        }
        return enterprises;
    }

    /*
     * 使用c3p0进行一个数据库写入
     * */
    public static Statement writeDataBase() throws PropertyVetoException, SQLException {
        comboPooledDataSource = new ComboPooledDataSource();
        comboPooledDataSource.setDriverClass("org.mariadb.jdbc.Driver");
        comboPooledDataSource.setJdbcUrl("jdbc:mariadb://localhost:3306/policy");
        comboPooledDataSource.setUser("root");
        comboPooledDataSource.setPassword("love1314");
        Connection connection = comboPooledDataSource.getConnection();
        Statement statement = connection.createStatement();
        return statement;
    }
}
