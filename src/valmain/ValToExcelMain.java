package valmain;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.time.Duration;
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
import java.util.Properties;
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
import org.apache.xmlbeans.impl.store.QueryDelegate.QueryInterface;

import valtoexcel.ExcelPoi;
import valtoexcel.Mail;
import valtoexcel.Resume;
import valtoexcel.Val;
import valtoexcel.Val.StatusVal;

public class ValToExcelMain {
	public static final Logger logger = Logger.getGlobal();

	public static String dirForFinalFile;
	public static String dirOfJarParent; 
	public static File hist;

	public static void main(String[] args) throws IOException, SQLException, InvalidFormatException, URISyntaxException {
		LocalTime start= LocalTime.now();
		configLogger();

		String dstOfTemplate = configDir();
		File fileTemplate = new File(dstOfTemplate);

		List<HashMap<String, String>> listTresh = getTresholds(fileTemplate);

		LinkedHashSet<Val> set = null;

		InputStream inSval = ValToExcelMain.class.getResourceAsStream("/val.txt");
		String dstValtxt = configDirValTxt(inSval);

		try {
			set = loadValsFromFile(new File(dstValtxt));
		} catch (Exception e) {
			e.printStackTrace();
			logger.log(Level.SEVERE, e.getMessage(), e.getStackTrace());

			try {
				set = loadVal();
				logger.info("Load default Query from Program");
			} catch (Exception e2) {

				logger.log(Level.SEVERE, e2.getMessage());
				return;
			}

		}

		InputStream in=ValToExcelMain.class.getResourceAsStream("/sqlQuerys.sql");
		InputStreamReader iSr = new InputStreamReader(in);

		Properties prop = new Properties();
		InputStream inProp = ValToExcelMain.class.getResourceAsStream("/connectpro.properties");
		prop.load(inProp);


		String url = prop.getProperty("url");

		ExcelPoi.setQuerysForValFromFile(iSr , set);

		//		Scanner scan =  new Scanner(System.in);
		//
		//		boolean out = false;
		//		do {
		//			System.out.println("Continuar ? ");
		//			String choice  = scan.next();
		//
		//			if (choice.equalsIgnoreCase("s")) {
		//				out = true;
		//			}else if(choice.equalsIgnoreCase("n")){
		//				System.exit(1);
		//			}
		//
		//		} while (!out);
		//		scan.close();

		try(		Connection c = DriverManager.getConnection(url,prop);

				) 

		{

			boolean onlyOnce = false;
			for (Val val : set) {
				SortedMap<Integer, List<String>> map = new TreeMap<>();
				List<String> listHead = new ArrayList<>();
				try 
				(
						PreparedStatement st = c.prepareStatement(val.getQuery());

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
						int colCount = result.getMetaData().getColumnCount();
						val.setMaxCollumn(colCount);
						for (int i = 1; i <= colCount; i++) {
							String res = result.getMetaData().getColumnName(i);
							listHead.add(res);
						}

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
					val.setHeadNames(listHead);
					val.setMap(map);
					logger.info("ResultMap: "+val.getName()+" "+val.getMap());
					//System.out.println(val.getName() + " ");// +val.getMap().values());
				} catch (Exception e) {
					logger.log(Level.SEVERE,e.getMessage(), e);

				}

			}
		}catch (Exception e) {

			logger.log(Level.SEVERE,e.getMessage(), e.getStackTrace());

			return;
		}
		List<Resume> listFinalResume = checkStatus(set, listTresh);

		ExcelPoi.whiteMapValExel(set,fileTemplate);
		addToHistoricFile(listFinalResume);
		printMap(listFinalResume);
		
		mailDraft(listFinalResume);
		
		Duration diff = Duration.between(start, LocalTime.now());
		logger.fine("Dura��o: "+diff.toMinutes()+"m:"+diff.getSeconds()+"s");



	}

	private static  void mailDraft(List<?> listFinalResume) {
		
		String sub = (listFinalResume.isEmpty()?"Monits OK" :"Monits NOK");
		
		Mail mail = new Mail(sub, ExcelPoi.finalFile, ExcelPoi.dirImageEmail, listFinalResume);
		mail.mailGenarator();
	}

	private static LinkedHashSet<Val> loadValsFromFile(File file) throws IOException  {
		//FileReader frVal  = new FileReader(new File("D:\\FileEx\\Val.txt"));

		FileReader fr = new FileReader(file);
		BufferedReader brVal =  new BufferedReader(fr);
		String strVal = "";
		int count = 0;
		LinkedHashSet<Val> setVal =  new LinkedHashSet<>();
		while ((strVal=brVal.readLine())!=null) {
			if(count==0) {
				count++;
				continue;			
			}


			String [] strArr = strVal.split(",",4);
			for (int i = 0; i < strArr.length; i++) {
				if (i==3) {
					strArr[i] = strArr[i].replace(';', ' ');
					continue;
				}
				strArr[i]= strArr[i].replace('"', ' ');
				strArr[i] = strArr[i].replace(';', ' ');
				strArr[i] = strArr[i].replace(')', ' ');
				strArr[i] = strArr[i].replace('(', ' ');
				strArr[i] = strArr[i].replace('-', ' ');

			}
			String name  = strArr[0].trim();
			int line = Integer.parseInt(strArr[1].trim());
			int col = Integer.parseInt(strArr[2].trim());

			Val val = new Val(name, line, col);


			if(strArr.length>3 && strArr[3].toLowerCase().contains("select")) {
				val.setQuery(strArr[3]);
			}

			setVal.add(val);
		}

		if (setVal.isEmpty()) {

			logger.warning("Map vazio");
			brVal.close();
			return loadVal();
		}
		brVal.close();
		logger.info("Loaded Val from File");
		return setVal;
	}
	private static LinkedHashSet<Val> loadVal() {
		logger.info("Load predefined validations");
		Val val1 = new Val( "VAL1", 3 , 0,  "SELECT   * FROM job_history ", "VAL1: Processo N�o terminado com �ltima mensagem a NotValidated");
		Val val2 = new Val( "VAL2", 4, 0 );
		Val val2_1 = new Val( "VAL2.1", 4 , 5);
		Val val4 = new Val( "VAL4", 3, 0);
		Val val5 = new Val( "VAL5", 3 , 0);
		/* For testing */


		LinkedHashSet<Val> set = new LinkedHashSet<>();
		set.add(val1);
		set.add(val2);
		set.add(val2_1);
		set.add(val4);
		set.add(val5);
		return set;
	}
	private static String configDirValTxt(InputStream inSval) throws URISyntaxException, IOException {
		String dirTxt = ValToExcelMain.class.getProtectionDomain().getCodeSource().getLocation().getFile();

		String dirTxtPar =  new File(dirTxt).getParent();

		new File(dirTxtPar+"\\Monitoriza��es").mkdirs();

		String dst = dirTxtPar+"\\Monitoriza��es\\Val.txt";
		logger.info("Diretorio do txt " + dst);
		if (!new File(dst).exists()) {
			Files.copy(inSval, Paths.get(dst));

			logger.info("Txt criado " + new File(dst).getName());
		}else if(new File(dst).exists() && new File(dst).length()<1){
			Files.copy(inSval, Paths.get(dst), StandardCopyOption.REPLACE_EXISTING);
			logger.log(Level.CONFIG, "Tamanho invalido Val.txt");
		}

		return dst;
	}
	private static String configDir() throws IOException {


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
			logger.info("Template criado");
			Files.copy(inputS, Paths.get(dst) );
		}else if (new File(dst).exists()&&(new File(dst).length())<=2000) {

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

		LogFormatter ff = new LogFormatter();

		SimpleFormatter formaterFile = new SimpleFormatter() {

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


		ch.setFormatter(formaterFile);

		fh.setFormatter(formaterFile);
		logger.addHandler(ch);
		logger.addHandler(fh);
		logger.setLevel(Level.ALL);
		logger.fine("Validation To Excel Executed   "+ mesToLogger.getTime().toString());
	}
	private static List<HashMap<String, String>> getTresholds(File file) throws InvalidFormatException, IOException {


		@SuppressWarnings("resource")
		XSSFWorkbook work = new XSSFWorkbook(file);
		XSSFSheet sheet = work.getSheet("Legenda");
		List<XSSFTable> tables = sheet.getTables();
		List<HashMap<String, String> > tableListMap = new ArrayList<>();
		for (XSSFTable t : tables) {
			HashMap<String, String> treshMap = new HashMap<>();
			if(t.getDisplayName().equals("TblOkNok")){
				continue;
			}

			int start = t.getStartRowIndex()+1;
			int end = t.getEndRowIndex();

			for (int i = start; i <= end; i++) {
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

	private static List<Resume> checkStatus(LinkedHashSet<Val> setVal, List<HashMap<String, String>> listTresh) {
		List<Resume> listResumeNOK = new ArrayList<>();
		for (Val val : setVal) {
			Resume resume = new Resume();
			boolean status = true;
			if(val.getMap()==null || val.getMap().isEmpty()) {
				val.setStatus(StatusVal.OK);

				continue;
			}
			if (val.getName().equals("VAL2")||val.getName().equals("VAL2.1")||val.getName().equals("VAL1")) {
				int indexTres = (val.getName().equals("VAL2")||val.getName().equals("VAL2.1"))? 1: 0;

				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {

					String bo = entry.getValue().get(1);
					double nMsg = Double.parseDouble(entry.getValue().get(0));
					//System.out.println(nMsg);
					if (bo!=null &&listTresh.get(indexTres).containsKey(bo)) {


						if(Double.parseDouble(listTresh.get(indexTres).get(bo))<=nMsg) {
							//System.out.println(val.getName() + " NOK " + bo + " " + nMsg);
							String res = bo+ " " + nMsg +" | "; 
							if (resume.getValName()!=val.getName()) {
								resume.setValName(val.getName());
								resume.setResumeStrB(new StringBuilder(res));
								listResumeNOK.add(resume);

							}else {
								//listResumeNOK.get(val.getName()).append(res);
								resume.getResumeStrB().append(res);
							}

							status=false;
						}
					}
					else if(nMsg>=Double.parseDouble(listTresh.get(indexTres).get("Default"))) {
						double defaultValue = Double.parseDouble(listTresh.get(indexTres).get("Default"));
						System.out.println(val.getName() + " NOK " + bo + " " + nMsg + " Default value used "+defaultValue);
						String res = bo+ " " + nMsg +" | "; 
						if (resume.getValName()!=val.getName()) {
							resume.setValName(val.getName());
							resume.setResumeStrB(new StringBuilder(res));
							listResumeNOK.add(resume);
							//listResumeNOK.add(val.getName(), new StringBuilder(res));
						}else {
							resume.getResumeStrB().append(res);
							//listResumeNOK.get(val.getName()).append(res);
						}

						status=false;
					}
					val.setStatus((status ? StatusVal.OK : StatusVal.NOK));
				}

			}else if(!val.getName().equals("VAL9")) {
				int size =val.getMap().size();
				val.setStatus(StatusVal.NOK);
				List<String> adicionalInfo = new ArrayList<>();adicionalInfo.add(val.getHeadNames().get(0) + ": ");

				for (Entry<Integer, List<String>> entry : val.getMap().entrySet()) {

					adicionalInfo.add(entry.getValue().get(0));

				}
				resume.setTotalResult(size);
				resume.setValName(val.getName());
				resume.setResumeStrB(new StringBuilder(size +" Resultados " ));
				resume.setListValue(adicionalInfo);
				listResumeNOK.add(resume);
				//linkMapResumeNOK.put(val.getName(), new StringBuilder(size +" Resultados " ).append(adicionalInfo) );
			}
			if(val.getStatus()==null) {
				val.setStatus(StatusVal.OK);
			}


		}
		System.out.println(listResumeNOK);
		return listResumeNOK;

	}

	private static void printMap(List<Resume> listFinalResume) throws IOException {

		Iterator<Resume> it =listFinalResume.iterator();
		if(listFinalResume.isEmpty()) {
			logger.fine("-----------------------Resumo OK-----------------------");

		}else {
			logger.warning("-----------------------Resumo NOK-----------------------");

			while (it.hasNext()) {
				Resume res =it.next();
				String checkListValues = checkIfAlreadyExistsInHistoriy(res.getListValue(),hist);
				logger.warning(res.toString() + checkListValues);

			}
		}


	}

	private static void addToHistoricFile(List<Resume> listFinalResume) throws IOException {
		List<String> ignoreList =  Arrays.asList("VAL1","VAL2", "VAL2.1","VAL14");
		new File(dirOfJarParent+"\\Historico").mkdirs();
		hist = new File(dirOfJarParent+"\\Historico\\historico.log");
		
		Calendar today = Calendar.getInstance();
		//today.add(Calendar.DAY_OF_MONTH, 16);

		SimpleDateFormat formart = new SimpleDateFormat("dd-MM-yyyy");
		String todayString = formart.format(today.getTime());
		
		try (
				FileWriter fw = new FileWriter(hist,true);
				BufferedReader br = new BufferedReader(new FileReader(hist));

				)
		{

			String brLine="";
			boolean alreadyAddedToday= false;
			while ((brLine=(br.readLine()))!=null) {
				if(brLine.contains(todayString)) {
					alreadyAddedToday= true;
					logger.info("J� adicinado no historico");
				}

			}
			if(!alreadyAddedToday) {
				logger.info("Add results to historic file "+ todayString+" " +hist.getName());
				Iterator<Resume>it = listFinalResume.iterator();
				fw.write("\n"+todayString+": ");
				while (it.hasNext()) {

					Resume next = it.next();
					if(!ignoreList.contains(next.getValName()) ) {

						fw.write(next.toString() + next.toStringList(next.getListValue()) +" | ");
					}


				}

			}
			//fw.write("\n");

		} catch (Exception e) {
			e.printStackTrace();
		}


	}
	private static String checkIfAlreadyExistsInHistoriy(List<String> list , File file) throws IOException {
		//logger.info("Check in history file");
		StringBuilder strBuild = new StringBuilder();
		
		for (String string : list) {
			int count = 0;
			if (string.equals(list.get(0))) {
				strBuild.append(string+" ");
				continue;
			}
			try(
					BufferedReader br = new BufferedReader(new FileReader(file));

					) 
			{
				String line;
				while ((line=br.readLine())!=null) {
					if (line.contains(string)) {
						
						count++;
					}

				}	
			} catch (Exception e) {
				logger.severe(e.getLocalizedMessage());
			}

			switch (count) {
			case 1:
				
				strBuild.append(string+" (novo), ");
				
				break;
			case 2:
				strBuild.append(string+" (recente), ");break;
			default:
				strBuild.append(string+", ");break;
			}

		}
		//System.out.println(strBuild);
		return strBuild.toString();
	}
}





