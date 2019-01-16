package valtoexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map.Entry;
import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface ExcelPoi {
	


	//Create Blank Workbook

	public static void whiteMapValExel(SortedMap<Integer, List<String>> map) throws IOException {

		File file = new File("D:\\FileEx\\Livro1.xlsx");
		try(
				FileInputStream fi = new FileInputStream(file);

				XSSFWorkbook work = new XSSFWorkbook(fi);

				)
		{
			XSSFSheet sheet = work.getSheet("VAL1");
			int line =1; 
			for (Entry<Integer, List<String>> entry : map.entrySet()) {

				XSSFRow row = sheet.createRow(line);
				int collumnCell = 1;
				for (String row2 : entry.getValue()) {

					XSSFCell cell = row.createCell(collumnCell);
					cell.setCellValue(row2);
					if(row2.matches("[0-9]+")){
						cell.setCellType(CellType.NUMERIC);
						//System.out.println(row2+ " numeic");

					}
					collumnCell++;

				}
				line++;
			}
			FileOutputStream out = new FileOutputStream(file);

			work.write(out);
			System.out.println("ficheiro Criado: " + file.getName());
		}



	}
	/**
	 * 
	 * 
	 * @param HashSet of Sql results of Valations
	 * @author Ricardo Russo
	 * @throws IOException
	 */
	public static void whiteMapValExel(Set<Val> set) throws IOException {
		Calendar c = Calendar.getInstance();
		SimpleDateFormat format = new SimpleDateFormat("dd-MM-yyyy k mm ", Locale.getDefault());
		SimpleDateFormat formatMes = new SimpleDateFormat("MMMMM", Locale.getDefault());

		String date =  format.format(c.getTime());
		String mes = formatMes.format(c.getTime());
		System.out.println("Write to excel");

		File file = new File("D:\\FileEx\\Livro1"+".xlsx");
		
		try(
				FileInputStream fi = new FileInputStream(file);

				XSSFWorkbook work = new XSSFWorkbook(fi);

				)
		{
			for (Val val : set) {
				XSSFSheet sheet = work.getSheet(val.getName());
				if (sheet==null) {
					sheet = work.createSheet(val.getName());
				}
				int line =val.getLine(); 
				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {

					XSSFRow row = sheet.createRow(line);
					int collumnCell = val.getCol();
					for (String row2 : entry.getValue()) {

						XSSFCell cell = row.createCell(collumnCell);
						cell.setCellValue(row2);
//						if(row2.matches("[0-9]+")){
//							cell.setCellType(CellType.NUMERIC);
//							//System.out.println(row2+ " numeric");
//
//						}
						collumnCell++;

					}
					line++;
				}

			}
			
			File newFile = new File("D:\\FileEx\\Livro1_"+ date +".xlsx");
			
			//newFile.createNewFile();
			FileOutputStream out = new FileOutputStream(newFile);
			
			work.write(out);
			System.out.println("ficheiro Criado: " + newFile.getName());
		}

	}
	
}
