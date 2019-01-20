package valtoexcel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import java.util.Locale;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


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
	public static void whiteMapValExel(Set<Val> set) throws IOException, EncryptedDocumentException, InvalidFormatException {
		Calendar c = Calendar.getInstance();
		SimpleDateFormat format = new SimpleDateFormat("dd-MM-yyyy k mm ", Locale.getDefault());

		String date =  format.format(c.getTime());
		//String mes = formatMes.format(c.getTime());
		//SimpleDateFormat formatMes = new SimpleDateFormat("MMMMM", Locale.getDefault());

		System.out.println("Write to excel");

		File file = new File("D:\\FileEx\\Livro1"+".xlsx");


		try(
				FileInputStream fi = new FileInputStream(file);

				Workbook work = WorkbookFactory.create(fi);

				)
		{
			for (Val val : set) {
				Sheet sheet = work.getSheet(val.getName());
				if (sheet==null) {
					sheet = work.createSheet(val.getName());
				}
				int line =val.getLine(); 
				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {

					Row row = sheet.createRow(line);
					int collumnCell = val.getCol();
					for (String row2 : entry.getValue()) {

						Cell cell = row.createCell(collumnCell);
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
					// TODO: handle exception
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

}

