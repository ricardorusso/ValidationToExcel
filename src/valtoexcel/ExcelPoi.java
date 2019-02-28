package valtoexcel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.Array;
import java.nio.charset.StandardCharsets;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.format.CellNumberFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;


public  abstract class ExcelPoi {

	/**
	 * 
	 * 
	 * @param HashSet of Sql results of Valations
	 * @author Ricardo Russo
	 * @throws IOException
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	public static void whiteMapValExel(Set<Val> set, File template) throws IOException, EncryptedDocumentException, InvalidFormatException {
		Calendar c = Calendar.getInstance();
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd k mm ", Locale.getDefault());
		SimpleDateFormat formatLastMonit = new SimpleDateFormat("dd/MM");
		String date =  format.format(c.getTime());
		
		System.out.println("Write to excel");

		try(
				FileInputStream fi = new FileInputStream(template);

				Workbook work = WorkbookFactory.create(fi);

				)
		{
	
		
			for (Val val : set) {
				Sheet sheet = work.getSheet(val.getName());
				if(val.getName().equals("VAL2.1")) {
					sheet = work.getSheet("VAL2");
				}
				if (sheet==null) {
					sheet = work.createSheet(val.getName());
				}
				
				int line =val.getLine(); 
				if(val.getMap()== null ) {
					System.err.println("Erro "+val.getName() + " map null");
					continue;
				}
				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {

					Row row = sheet.getRow(line);
					if (row == null) {
						row = sheet.createRow(line);
					}
					//System.out.println(line);
					int collumnCell = val.getCol();
					for (String row2 : entry.getValue()) {

						Cell cell = row.getCell(collumnCell);
						if (cell == null) {
							cell= row.createCell(collumnCell);
						}
						
						if( row2!=null && row2.matches("^[0-9]*$") ){
							cell.setCellValue(Long.parseLong(row2));
							
						
							//System.out.println(row2+ " numeric");
							
						}else {
							cell.setCellValue(row2);
						}
						collumnCell++;
					}
					line++;
				}

			}
			
			File newFile = new File(template.getPath()+"_"+ date +".xlsx");
			
			addTimeline(formatLastMonit, work);
			
			FileOutputStream out = new FileOutputStream(newFile);
			work.setForceFormulaRecalculation(true);
			work.write(out);
			System.out.println("ficheiro Criado: " + newFile.getName());
		}

	}
	private static void addTimeline(SimpleDateFormat formatLastMonit, Workbook work) {
		System.out.println("Add timeline");
		Calendar c = Calendar.getInstance();
		Calendar lastMonitDate = Calendar.getInstance();
		lastMonitDate.add(Calendar.DATE, -1);
		
		switch (lastMonitDate.get(Calendar.DAY_OF_WEEK)) {
		case 1:
			lastMonitDate.add(Calendar.DATE, -2);
			break;
		case 7:
			lastMonitDate.add(Calendar.DATE, -1);
			break;
		default:
			break;
		}
		System.out.println(lastMonitDate.getTime());

		Sheet main = work.getSheet("Overview");
		Row row = main.getRow(2);
		Cell cell = row.getCell(1);
		cell.setCellValue(lastMonitDate.getTime());
		//newFile.createNewFile();
		
		Sheet timeline = work.getSheet("Timeline");
		Row rowTimeLine = timeline.getRow(timeline.getLastRowNum()); //72
		System.out.println(rowTimeLine.getRowNum());//16

		int lastRow = timeline.getLastRowNum();
		List<String> listOkVal = new ArrayList<>(Arrays.asList("'VAL1'!F1", "'VAL2'!E1","'VAL3'!F1","'VAL4'!L1","'VAL5'!M1","'VAL6'!M1","'VAL7'!M1","'VAL8'!M1","'VAL9'!I1","'VAL10'!I1","'VAL11'!I1","'VAL12'!K1"));
		
		int lastCell = rowTimeLine.getLastCellNum();
		rowTimeLine = timeline.getRow(lastRow-12);
		Cell headTime = rowTimeLine.createCell(lastCell);
		headTime.setCellValue(formatLastMonit.format(c.getTime()));
		
		int countListOk = 0;
		for (int i = (lastRow-11); i <= lastRow; i++) {
			rowTimeLine = timeline.getRow(i);
			Cell cellTime= rowTimeLine.createCell(lastCell);
			cellTime.setCellFormula(listOkVal.get(countListOk));
			countListOk++;
			
		}
		
	}
	public static File fileXml = new File("D:\\FileEx\\TABLE_EXPORT_DATA_2.xml");

	public static void splitFileContentToSeperateFiles(File xml, String begin, String end) {
		if(!xml.isFile()) {
			System.out.println("Not a file");
			return;
		}
		try (			
				FileReader fr= new FileReader(xml);
				BufferedReader br= new BufferedReader(fr);		
				)
		{
			StringBuilder strB = new StringBuilder();
			String ln = "";
			while ((ln = br.readLine())!=null) {
				strB.append(ln);				
			}
			//System.out.println(strB);
			int nameCount = 1;
			while (strB.indexOf(begin)!=-1) {
				int row = strB.indexOf(begin);
				int rowEnd = strB.indexOf(end);
				String subString = strB.substring(row+(begin.length()),rowEnd);
				File newFile = new File("D:\\FileEx\\Nova\\"+nameCount+"_.xml");
				FileWriter fw = null;
				try 
				{
					fw = new FileWriter(newFile);
					fw.write(subString);
					System.out.println("Ficheiro Criado " +newFile.getName()+" em "+newFile.getPath());
				} catch (Exception e) {
					e.printStackTrace();
				}finally {
					if(fw!=(null))fw.close();
				}
				nameCount++;
				System.out.println();
				//System.out.println(newS);
				strB = new StringBuilder( strB.substring(rowEnd+1, strB.length()));
				//System.out.println(strB);
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}
	public static void setQuerysForValFromFile(File sql, Set<Val> set) throws IOException {

		try(
				BufferedReader br = new BufferedReader(new FileReader(sql));
				){
			String str = "";
			StringBuilder strBu = new StringBuilder();
			while (((str=br.readLine())!=null)) {
				strBu.append(str+" " );

			}
			for (Val val : set) {
				int start =  strBu.indexOf(val.getName());
				int end = strBu.indexOf(";", start);
				String substring =  strBu.substring(start+val.getName().length(), end);

				String cleanString = substring.replace('-', ' ').trim();
				if (val.getQuery()==null ) {
					val.setQuery(cleanString);
					System.out.println(val.getName() + " "+cleanString+"\n");
					

				}else {
					System.out.println(val.getName() + " Query already set "+ val.getQuery());
				}
				
			}
		}catch (Exception e) {
			e.printStackTrace();
		}
	}

}

