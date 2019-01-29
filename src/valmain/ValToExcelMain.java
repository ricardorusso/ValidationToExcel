package valmain;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import valtoexcel.Constants;
import valtoexcel.ExcelPoi;
import valtoexcel.Val;

public class ValToExcelMain {

	public static void main(String[] args) throws IOException, SQLException, EncryptedDocumentException, InvalidFormatException {

		//		connection 
		String user ="NOVO";
		String pass = "novo";
		String url = "jdbc:oracle:thin:@//localhost:1521/xe";

		Val val1 = new Val( "VAL1", 2 , 5, 2);
		Val val2 = new Val( "VAL2", 2, 5, 4);
		Val val3 = new Val( "VAL3", 2 , 5, 2);
		Val val4 = new Val( "VAL4", 2, 5, 2);
		Val val5 = new Val( "VAL5", 2 , 5, 4);

		LinkedHashSet<Val> set = new LinkedHashSet<>();
		set.add(val1);
		set.add(val2);
		set.add(val3);
		set.add(val4);
		set.add(val5);

		ExcelPoi.setQuerysForValFromFile(new File("D:\\FileEx\\sqlQuerys.sql"),set);
		Connection c = DriverManager.getConnection(url, user, pass);
		for (Val val : set) {

			SortedMap<Integer, List<String>> map = new TreeMap<>();
			try 
			(
					Statement st = c.prepareStatement(val.getQuery());
					)
			{				c.setReadOnly(true);

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

						} catch (SQLException e) {
							System.out.println("Coluna " +i +" Não existe ");
							val.setMaxCollumn(i-1);
							break resultSql;
						}				
					}

					map.put(line,list2);
				}

			}
			val.setMap(map);
			System.out.println("MAp: "+val.getMap());
			//System.out.println(val.getName() + " ");// +val.getMap().values());
			} catch (Exception e) {
				// TODO: handle exception
			}

		}

		File template = new File("D:\\FileEx\\Livro1"+".xlsx");
		ExcelPoi.whiteMapValExel(set,template);



	}

}


