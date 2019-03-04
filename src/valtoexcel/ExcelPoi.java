package valtoexcel;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Locale;
import java.util.Map.Entry;
import java.util.Set;
import java.util.logging.Logger;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import valmain.ValToExcelMain;
import valtoexcel.Val.StatusVal;


public  abstract class ExcelPoi {
	private static final Logger logger = ValToExcelMain.logger;
	private static final String VAL2_1 = "VAL2.1";
	/**
	 * 
	 * 
	 * @param HashSet of Sql results of Valations
	 * @author Ricardo Russo
	 * @throws IOException
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	public static void whiteMapValExel(Set<Val> set, File template) throws IOException, InvalidFormatException {
		
		
		Calendar c = Calendar.getInstance();
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd k mm ", Locale.getDefault());

		String date =  format.format(c.getTime());
		List<String> listOkVal = new ArrayList<>();
		System.out.println("-----------------Write to excel-----------------");

		try(
				FileInputStream fi = new FileInputStream(template);

				Workbook work = WorkbookFactory.create(fi);

				)
		{


			for (Val val : set) {
				//listOkVal.add((val.getStatus().equals(StatusVal.OK)?"OK":"NOK"));
				if(val.getName().equals(VAL2_1)&&val.getStatus().equals(StatusVal.NOK)&&listOkVal.get(listOkVal.size()-1).equals("OK")) {
					listOkVal.remove(listOkVal.size()-1);
					listOkVal.add(val.getStatus().getOkNok());
					
				}else if(!val.getName().equals(VAL2_1)){
					listOkVal.add(val.getStatus().getOkNok());
				}
				


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

			try {
				addTimeline(work, listOkVal);

			} catch (Exception e) {
				e.printStackTrace();
			}

			FileOutputStream out = new FileOutputStream(newFile);
			work.setForceFormulaRecalculation(true);
			work.write(out);
			System.out.println("ficheiro Criado: " + newFile.getName());
		}

	}
	private static void addTimeline( Workbook work, List<String> listOkVal) {
		//System.out.println(listOkVal);
		System.out.println("-----------------Add timeline-----------------");
		SimpleDateFormat formatLastMonit = new SimpleDateFormat("dd/MMM");
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
		
		int lastCell = rowTimeLine.getLastCellNum();
		rowTimeLine = timeline.getRow(lastRow-12);

		Cell headTime = rowTimeLine.createCell(lastCell);
		headTime.setCellValue(formatLastMonit.format(c.getTime()));

		int countListOk = 0;
		for (int i = (lastRow-11); i <= lastRow; i++) {
			if(countListOk>=listOkVal.size()) {
				return;
			}
			rowTimeLine = timeline.getRow(i);
			Cell cellTime= rowTimeLine.createCell(lastCell);

			cellTime.setCellValue(listOkVal.get(countListOk));
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
		if(!sql.exists()) {
			System.err.println("Sql ficheiro n�o existe ");
			return;
		}
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

