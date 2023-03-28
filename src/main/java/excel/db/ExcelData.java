package excel.db;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class ExcelData {

	final static String url = "jdbc:mysql://localhost:3306/raju";
	final static String user_host = "root";
	final static String pwd = "141199";
	private static final String select = " select * from filter";

	public static void main(String[] args) throws Exception {

		Connection con = DriverManager.getConnection(url, user_host, pwd);
		Statement stmt = con.createStatement();
		ResultSet rs = stmt.executeQuery(select);
		System.out.println("data fetched");

		// Create Excel workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Excel");

		// Write the column headers to the first row of the worksheet
		Row headerRow = sheet.createRow(0);
		int columnCount = rs.getMetaData().getColumnCount();
		for (int i = 1; i <= columnCount; i++) {
			Cell cell = headerRow.createCell(i - 1);
			cell.setCellValue(rs.getMetaData().getColumnName(i));
		}

		// Write data to sheet
		int rowNum = 1;
		while (rs.next()) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(rs.getInt(1));
			row.createCell(1).setCellValue(rs.getString(2));

			row.createCell(2).setCellValue(rs.getString(3));
			System.out.println("inserting");
			row.createCell(3).setCellValue(rs.getString(4));
			System.out.println("inserting");
			row.createCell(4).setCellValue(rs.getString(5));
			System.out.println("inserting");
			row.createCell(5).setCellValue(rs.getInt(6));
			System.out.println("inserting");

		}

		// Save workbook to file
		File file = new File("D:\\imp\\Excel.xlsx");
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();
		System.out.println("EmployeeData.xlsx written successfully!");

	}
}
