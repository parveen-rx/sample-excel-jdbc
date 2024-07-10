import java.sql.*;

public class Main {
    public static void main(String[] args) {
        Connection connection = null;
        try {
            Class.forName("com.hxtt.sql.excel.ExcelDriver").getDeclaredConstructor().newInstance();
            String excelFilePath = "D:\\test.xlsx";
            String url = "jdbc:Excel:////" + excelFilePath;
            connection = DriverManager.getConnection(url, "", "");
            Statement statement = connection.createStatement();
            String selectQuery = "SELECT * FROM Sheet1";
            ResultSet resultSet = statement.executeQuery(selectQuery);
            while (resultSet.next()) {
                System.out.println(resultSet.getString(1) + "\t" + resultSet.getString(2));
            }

            // Example INSERT query
            String insertQuery = "INSERT INTO Sheet1 (Id, Username) VALUES (6, 'custom-admin')";
            statement.executeUpdate(insertQuery);
            statement.close();
            System.out.println("Sheet 1 update is complete..");
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(connection != null){
                try {
                    connection.close();
                } catch (SQLException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}