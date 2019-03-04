package valmain;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map.Entry;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.logging.ConsoleHandler;
import java.util.logging.FileHandler;
import java.util.logging.Filter;
import java.util.logging.Formatter;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.LogRecord;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import javax.net.ssl.SSLEngineResult.Status;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import valtoexcel.Constants;
import valtoexcel.ExcelPoi;
import valtoexcel.Val;
import valtoexcel.Val.StatusVal;

public class ValToExcelMain {
	public static final Logger logger = Logger.getGlobal();
	private static File template = new File("D:\\FileEx\\MonitorCSW_v12_unres1"+".xlsx");
	private static File template2 = new File("C:\\Users\\Ricardo Russo\\Google Drive\\Ficheiros Empresas\\Accenture\\MonitorCSW_v12_unres1.xlsx");
	private static File fileSql = new File("D:\\FileEx\\sqlQuerys.sql");
	public static void main(String[] args) throws IOException, SQLException, InvalidFormatException {
	
		configLogger();
		
		
		
		List<HashMap<String, String>> listTresh = getTresholds(template2);
		
		SortedMap<Integer, List<String>> mapExempleVal = new TreeMap<>();
		mapExempleVal.put(1, Arrays.asList("5000.0","CM_PT_UPLREADS"));
		mapExempleVal.put(2, Arrays.asList("5000.0","CM_SPA_E"));
		mapExempleVal.put(3, Arrays.asList("5000.0","CM_SPA_E_SWINOT"));
		mapExempleVal.put(4, Arrays.asList("5000.0","CM_POR_G_H1_N6005"));
		
		LocalTime start= LocalTime.now();

		//		connection 
		String user ="NOVO";
		String pass = "novo";
		String url = "jdbc:oracle:thin:@//localhost:1521/xe";

		Val val1 = new Val( "VAL1", 2 , 5, 2, "SELECT   /*+ PARALLEL(16)*/  TO_DATE(end_date) Data FROM job_history where end_date = '06.07.24'");
		Val val2 = new Val( "VAL2", 2, 10, 45);
		Val val3 = new Val( "VAL2.1", 2 , 5, 2);
		Val val4 = new Val( "VAL4", 2, 5, 2);
		Val val5 = new Val( "VAL5", 2 , 5, 10);
		/* For testing */
		val1.setMap(mapExempleVal);
		val2.setMap(mapExempleVal);
		val3.setMap(mapExempleVal);
		val4.setMap(mapExempleVal);
		val5.setMap(new TreeMap<>());
		
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
						System.out.println(val.getName() + " -----------------Query Executed-----------------");
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
										System.out.println("Coluna " +i +" Não existe ");
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
					logger.log(Level.SEVERE,e.getMessage(), e.getStackTrace());
				}

			}
		}catch (Exception e) {
		
			logger.log(Level.SEVERE,e.getMessage(), e.getStackTrace());
		}
		LinkedHashMap<String, StringBuilder> linkMapResumeNOK = checkStatus(set, listTresh);
		ExcelPoi.whiteMapValExel(set,template2);

		printMap(linkMapResumeNOK);
		Duration diff = Duration.between(start, LocalTime.now());
		System.out.println("Duração: "+diff.toMinutes()+":"+diff.getSeconds()+"s");
		
		

	}
	private static void configLogger() throws IOException {
		Calendar mesToLogger = Calendar.getInstance();
		SimpleDateFormat fDate = new SimpleDateFormat("MMMM");
		String mounth = fDate.format(mesToLogger.getTime());
		
		
	
		LogManager.getLogManager().reset();
		FileHandler fh = new FileHandler("ValToExel_"+mounth+".log",true);
		ConsoleHandler ch = new ConsoleHandler();
		ch.setLevel(Level.ALL);
		
		SimpleFormatter ff = new SimpleFormatter() {
			  private static final String FORMAT = "[%1$tF %1$tT] [%2$-7s] %3$s %n";

	          @Override
	          public synchronized String format(LogRecord lr) {
	              return String.format(FORMAT,
	                      new Date(lr.getMillis()),
	                      lr.getLevel().getLocalizedName(),
	                      lr.getMessage()
	              );
	          }
		};
		ch.setFormatter(ff);
		fh.setFormatter(ff);
		logger.addHandler(ch);
		logger.addHandler(fh);
		logger.setLevel(Level.ALL);
		logger.info("Validation To Excel Executed   "+mesToLogger.getTime().toString());
	}
	private static List<HashMap<String, String>> getTresholds(File file) throws InvalidFormatException, IOException {

		Workbook work = WorkbookFactory.create(file);
		XSSFSheet sheet = (XSSFSheet) work.getSheet("Legenda");
		List<XSSFTable> tables = sheet.getTables();
		List<HashMap<String, String> > tableListMap = new ArrayList<>();
		for (XSSFTable t : tables) {
			HashMap<String, String> treshMap = new HashMap<>();
			if(t.getDisplayName().equals("TblOkNok")){
				continue;
			}

			int star = t.getStartRowIndex()+1;
			int end = t.getEndRowIndex();

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
				treshMap.put(valueCell1.trim(), valueCell2.trim());

			}
			tableListMap.add(treshMap);
			

		}
		work.close();
		return tableListMap;

	}

	private static LinkedHashMap<String, StringBuilder> checkStatus(LinkedHashSet<Val> setVal, List<HashMap<String, String>> listTresh) {
		LinkedHashMap<String, StringBuilder> linkMapResumeNOK = new LinkedHashMap<>();
		for (Val val : setVal) {
		
			boolean status = true;
			if(val.getMap().isEmpty()) {
				val.setStatus(StatusVal.OK);
				
				continue;
			}
			if (val.getName().equals("VAL2")||val.getName().equals("VAL2.1")||val.getName().equals("VAL1")) {
				int indexTres = (val.getName().equals("VAL2")||val.getName().equals("VAL2.1"))? 1: 0;

				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {
					double nMsg = Double.parseDouble(entry.getValue().get(0));
					String bo = entry.getValue().get(1);
			
					if (listTresh.get(indexTres).containsKey(bo)) {
						if(Double.parseDouble(listTresh.get(indexTres).get(bo))<=nMsg) {
							System.out.println(val.getName() + " NOK " + bo + " " + nMsg);
							String res = bo+ " " + nMsg +" | "; 
							if (!linkMapResumeNOK.containsKey(val.getName())) {
								linkMapResumeNOK.put(val.getName(), new StringBuilder(res));
							}else {
								linkMapResumeNOK.get(val.getName()).append(res);
							}
							
							status=false;
						}
					}else if(nMsg>=Double.parseDouble(listTresh.get(indexTres).get("Default"))) {
						double defaultValue = Double.parseDouble(listTresh.get(indexTres).get("Default"));
						System.out.println(val.getName() + " NOK " + bo + " " + nMsg + " Default value used "+defaultValue);
						String res = bo+ " " + nMsg +" | "; 
						if (!linkMapResumeNOK.containsKey(val.getName())) {
							linkMapResumeNOK.put(val.getName(), new StringBuilder(res));
						}else {
							linkMapResumeNOK.get(val.getName()).append(res);
						}
						status=false;
					}
					val.setStatus((status ? StatusVal.OK : StatusVal.NOK));
				}
			}else if(!val.getName().equals("VAL9")) {
				int size =val.getMap().size();
				val.setStatus(StatusVal.NOK);
				linkMapResumeNOK.put(val.getName(), new StringBuilder(size +" Resultados"));
			}
			if(val.getStatus()==null) {
				val.setStatus(StatusVal.OK);
			}
			
			
		}
		
		return linkMapResumeNOK;

	}
	
	private static void printMap(LinkedHashMap<String, StringBuilder> linkMapResumeNOK) {
		
		Iterator<Entry<String, StringBuilder>> it =linkMapResumeNOK.entrySet().iterator();
		logger.info("-----------------------Resumo NOK-----------------------");
		
		while (it.hasNext()) {
			Entry<String, StringBuilder> entry =it.next();
			logger.info(entry.getKey() +" "+entry.getValue() );
			
		}
		
	}
}




