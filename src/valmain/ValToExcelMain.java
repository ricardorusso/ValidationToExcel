package valmain;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalTime;
import java.util.ArrayList;
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
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.LogRecord;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import valtoexcel.ExcelPoi;
import valtoexcel.Val;
import valtoexcel.Val.StatusVal;

public class ValToExcelMain {
	public static final Logger logger = Logger.getGlobal();
	//private static File template = new File("D:\\FileEx\\MonitorCSW_v12_unres1"+".xlsx");
	// static File template2 = new File("C:\\Users\\Ricardo Russo\\Google Drive\\Ficheiros Empresas\\Accenture\\MonitorCSW_v12_unres1.xlsx");
	//private static File fileSql = new File("D:\\FileEx\\sqlQuerys.sql");
	public static String dirForFinalFile;

	public static void main(String[] args) throws IOException, SQLException, InvalidFormatException {
		LocalTime start= LocalTime.now();
		configLogger();

		String dstOfTemplate = configDir();
		
		

		File fileTemplate = new File(dstOfTemplate);
		List<HashMap<String, String>> listTresh = getTresholds(fileTemplate);


		//		connection 
		String user ="NOVO";
		String pass = "novo";
		String url = "jdbc:oracle:thin:@//localhost:1521/xe";

		Val val1 = new Val( "VAL1", 3 , 0, 2, "SELECT   * FROM job_history ", "VAL1: Processo N�o terminado com �ltima mensagem a NotValidated");
		Val val2 = new Val( "VAL2", 4, 0, 3);
		Val val2_1 = new Val( "VAL2.1", 4 , 5, 2);
		Val val4 = new Val( "VAL4", 3, 0, 2);
		Val val5 = new Val( "VAL5", 3 , 0, 3);
		/* For testing */


		LinkedHashSet<Val> set = new LinkedHashSet<>();
		set.add(val1);
		set.add(val2);
		set.add(val2_1);
		set.add(val4);
		set.add(val5);


		InputStream in=ValToExcelMain.class.getResourceAsStream("/sqlQuerys.sql");
		InputStreamReader iSr = new InputStreamReader(in);

		ExcelPoi.setQuerysForValFromFile(iSr , set);

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
						logger.info("-----------------Connected---------------"+ c.getMetaData().getUserName()  );
						onlyOnce=true;
					}


					try (
							ResultSet result =  st.executeQuery(val.getQuery());

							){
						logger.info(val.getName() + " -----------------Query Executed-----------------");
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
										logger.warning("Coluna " +i +" N�o existe ");
										val.setMaxCollumn(i-1);
										break resultSql;
									}				
								}
							//System.out.println(list2);
							map.put(line,list2);
						}

					}

					val.setMap(map);
					logger.info("ResultMap: "+val.getName()+" "+val.getMap());
					//System.out.println(val.getName() + " ");// +val.getMap().values());
				} catch (Exception e) {
					logger.log(Level.SEVERE,e.getMessage(), e.getStackTrace());
					
				}

			}
		}catch (Exception e) {
			
			logger.log(Level.SEVERE,e.getMessage(), e.getStackTrace());
		
			return;
		}
		LinkedHashMap<String, StringBuilder> linkMapResumeNOK = checkStatus(set, listTresh);
		ExcelPoi.whiteMapValExel(set,fileTemplate);

		printMap(linkMapResumeNOK);
		Duration diff = Duration.between(start, LocalTime.now());
		logger.info("Dura��o: "+diff.toMinutes()+"m:"+diff.getSeconds()+"s");



	}
	private static String configDir() throws IOException {
		String dirOfJarParent = "";

		try {
			dirOfJarParent = ValToExcelMain.class.getProtectionDomain().getCodeSource().getLocation()
					.toURI().getPath();
		} catch (URISyntaxException e1) {

			logger.log(Level.SEVERE, e1.getMessage(), e1.getCause());
		}

		dirOfJarParent = new File(dirOfJarParent).getParent();
		
		//create Directory//
		Calendar mesForDir = Calendar.getInstance();
		SimpleDateFormat formatMes = new SimpleDateFormat("MMMM");
		new File(dirOfJarParent+"\\Monitoriza��es\\"+formatMes.format(mesForDir.getTime())).mkdirs();
		dirForFinalFile = dirOfJarParent+"\\Monitoriza��es\\"+formatMes.format(mesForDir.getTime());
		String dirForTemplate = dirOfJarParent+"\\Monitoriza��es";
		logger.info("Directorio do Template " +dirForTemplate);
		
		InputStream inputS= ValToExcelMain.class.getResourceAsStream("/Template.xlsx");

		String dst = dirForTemplate+"\\Template.xlsx";
		if(!new File(dst).exists()) {
			System.out.println("Template criado");
			Files.copy(inputS, Paths.get(dst) );
		}else if (new File(dst).exists()&&(new File(dst).length())<=100) {

			Files.copy(inputS, Paths.get(dst) , StandardCopyOption.REPLACE_EXISTING);
			logger.log(Level.CONFIG,"Tamanho invalido, template replaced");
		}
		return dst;
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
		logger.fine("Validation To Excel Executed   "+ mesToLogger.getTime().toString());
	}
	private static List<HashMap<String, String>> getTresholds(File file) throws InvalidFormatException, IOException {

		
		XSSFWorkbook work = new XSSFWorkbook(file);
		XSSFSheet sheet = work.getSheet("Legenda");
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
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell1 = row.getCell(t.getStartColIndex());
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
				XSSFCell cell2 = row.getCell(t.getEndColIndex());
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
			//work.close();
		}
		logger.info("Tresholds list added");
		return tableListMap;

	}

	private static LinkedHashMap<String, StringBuilder> checkStatus(LinkedHashSet<Val> setVal, List<HashMap<String, String>> listTresh) {
		LinkedHashMap<String, StringBuilder> linkMapResumeNOK = new LinkedHashMap<>();
		for (Val val : setVal) {

			boolean status = true;
			if(val.getMap()==null || val.getMap().isEmpty()) {
				val.setStatus(StatusVal.OK);

				continue;
			}
			if (val.getName().equals("VAL2")||val.getName().equals("VAL2.1")||val.getName().equals("VAL1")) {
				int indexTres = (val.getName().equals("VAL2")||val.getName().equals("VAL2.1"))? 1: 0;

				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {

					String bo = entry.getValue().get(1);

					if (bo!=null &&listTresh.get(indexTres).containsKey(bo)) {
						double nMsg = Double.parseDouble(entry.getValue().get(0));
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
						else if(nMsg>=Double.parseDouble(listTresh.get(indexTres).get("Default"))) {
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
		if(linkMapResumeNOK.isEmpty()) {
			logger.fine("-----------------------Resumo OK-----------------------");

		}else {
			logger.warning("-----------------------Resumo NOK-----------------------");

			while (it.hasNext()) {
				Entry<String, StringBuilder> entry =it.next();
				logger.warning(entry.getKey() +" "+entry.getValue() );

			}
		}


	}
}




