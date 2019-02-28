package valmain;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.Duration;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import valtoexcel.ExcelPoi;
import valtoexcel.Val;

public class ValToExcelMain {
	private static File template = new File("D:\\FileEx\\MonitorCSW_v12_unres1"+".xlsx");
	private static File fileSql = new File("D:\\FileEx\\sqlQuerys.sql");
	public static void main(String[] args) throws IOException, SQLException, InvalidFormatException {
		getTresholds(template);

		LocalTime start= LocalTime.now();

		//		connection 
		String user ="NOVO";
		String pass = "novo";
		String url = "jdbc:oracle:thin:@//localhost:1521/xe";

		Val val1 = new Val( "VAL1", 2 , 5, 2, "SELECT   /*+ PARALLEL(16)*/  TO_DATE(end_date) Data FROM job_history where end_date = '06.07.24'");
		Val val2 = new Val( "VAL2.1", 2, 10, 45);
		Val val3 = new Val( "VAL2", 2 , 5, 2);
		Val val4 = new Val( "VAL4", 2, 5, 2);
		Val val5 = new Val( "VAL5", 2 , 5, 10);

		LinkedHashSet<Val> set = new LinkedHashSet<>();
		set.add(val1);
		set.add(val2);
		set.add(val3);
		set.add(val4);
		set.add(val5);

		ExcelPoi.setQuerysForValFromFile(fileSql , set);

		try(		Connection c = DriverManager.getConnection(url, user, pass);
				) 

		{

			boolean onlyOnce = false;
			for (Val val : set) {

				SortedMap<Integer, List<String>> map = new TreeMap<>();
				try 
				(
						Statement st = c.prepareStatement(val.getQuery());
						)
				{				
					c.setReadOnly(true);
					
					if (!onlyOnce && !c.isClosed()) {
						System.out.println("-----------------Connected---------------"+ c.getMetaData().getURL()  );
						onlyOnce=true;
					}
					

					try (
							ResultSet result =  st.executeQuery(val.getQuery());

							){
						System.out.println(val.getName() + " Query Executed");
						int line = 0;
						while (result.next()) {
							line++;
							List<String> list2 = new ArrayList<>();

							resultSql:
								for (int i = 1; i <= val.getMaxCollumn(); i++) {
									try {
										list2.add(result.getString(i));
										//System.out.println(result.getString(i));
									} catch (SQLException e) {
										System.err.println("Coluna " +i +" Não existe ");
										val.setMaxCollumn(i-1);
										break resultSql;
									}				
								}
							//System.out.println(list2);
							map.put(line,list2);
						}

					}
					val.setMap(map);
					System.out.println("MAp: "+val.getMap());
					//System.out.println(val.getName() + " ");// +val.getMap().values());
				} catch (Exception e) {
					e.printStackTrace();
				}

			}
		}catch (Exception e) {
			e.printStackTrace();
		}

		ExcelPoi.whiteMapValExel(set,template);


		Duration diff = Duration.between(start, LocalTime.now());
		System.out.println("Duração: "+diff.toMinutes()+":"+diff.getSeconds()+"s");

	}
	private static void getTresholds(File file) throws EncryptedDocumentException, InvalidFormatException, IOException {

		Workbook work = WorkbookFactory.create(file);
		XSSFSheet sheet = (XSSFSheet) work.getSheet("Legenda");
		List<XSSFTable> tables = sheet.getTables();
		for (XSSFTable t : tables) {
			if(t.getDisplayName().equals("TblOkNok")){
				continue;
			}
			
			int star = t.getStartRowIndex()+1;
			int end = t.getEndRowIndex();
			LinkedHashMap<String, String> treshMap = new LinkedHashMap<>();
			for (int i = star; i <= end; i++) {
				Row row = sheet.getRow(i);
				Cell cell1 = row.getCell(t.getStartColIndex());
				String valueCell1 = "";
				String valueCell2 = "";
				if(cell1 != null) {
					
					switch (cell1.getCellTypeEnum()) {
					case NUMERIC :
						valueCell1 = Double.toString(cell1.getNumericCellValue());
						break;
					case STRING:
						valueCell1 = cell1.getStringCellValue();
						break;
					default:
						break;
					}
					
				}
				Cell cell2 = row.getCell(t.getEndColIndex());
				if(cell2 != null) {
					
					switch (cell2.getCellTypeEnum()) {
					case NUMERIC :
						valueCell2 = Double.toString(cell2.getNumericCellValue());
						break;
					case STRING:
						valueCell2 = cell2.getStringCellValue();
						break;
					default:
						break;
					}
					
				}
				treshMap.put(valueCell1, valueCell2);
				
			}
			System.out.println(treshMap);
			
		}


	}


}




