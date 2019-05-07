
import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainClass {
	public static void main(String[] args) throws Exception {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet(" Student Info ");
		XSSFRow row;

		FileOutputStream out = new FileOutputStream(new File("C:/poiexcel/Writesheet.xlsx"));
		
		
		String[] header = { "Student ID", "Student NAME", "points" };

		int cellid = 0;
		
		row = spreadsheet.createRow(0);
		
		for(String colnames : header)
		{
			XSSFCell cell = row.createCell(cellid++);
			cell.setCellValue(colnames);
		}
		
		String query ="Select name from student";
		
		List<Object> dbValues = getValuesFromDb(query);
	  int rowid =1;
		
		for(int i =0; i < dbValues.size();i++)
    {
			row = spreadsheet.createRow(rowid++);
			XSSFCell cell = row.createCell(1);
			cell.setCellValue((String)dbValues.get(i));
    }
		
		workbook.write(out);
		out.close();
		workbook.close();
		System.out.println("Writesheet.xlsx written successfully");
	}
	
	
	private static List<Object> getValuesFromDb(String query)
	{
		List<Object> dbValues = new ArrayList();
	  
		try {
			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/rest", "root", "system");

			Statement stmt = con.createStatement();
			ResultSet rs = stmt.executeQuery(query);
			while (rs.next()) {
				System.out.println((rs.getObject(1)));
			Object ob = rs.getObject(1);
			dbValues.add(ob);		
			}
			con.close();
		} catch (Exception e) {
			System.out.println(e);
		}
		
		return dbValues;
	}
}
