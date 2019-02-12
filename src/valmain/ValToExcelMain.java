package valmain;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.BreakIterator;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.record.PageBreakRecord.Break;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import valtoexcel.Constants;
import valtoexcel.ExcelPoi;
import valtoexcel.Val;

public class ValToExcelMain {
	private static File template = new File("D:\\FileEx\\Livro1"+".xlsx");
	private static File fileSql = new File("D:\\FileEx\\sqlQuerys.sql");
	public static void main(String[] args) throws IOException, SQLException, EncryptedDocumentException, InvalidFormatException {

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
//		set.add(val4);
//		set.add(val5);

		ExcelPoi.setQuerysForValFromFile(fileSql , set);
		Connection c = DriverManager.getConnection(url, user, pass);
		
		for (Val val : set) {

			SortedMap<Integer, List<String>> map = new TreeMap<>();
			try 
			(
					Statement st = c.prepareStatement(val.getQuery());
					)
			{				//c.setReadOnly(true);

			System.out.println("Connected"  );

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
				e.printStackTrace();
			}

		}
		ExcelPoi.whiteMapValExel(set,template);
	}

}


