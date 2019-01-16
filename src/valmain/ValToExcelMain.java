package valmain;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.SortedMap;
import java.util.TreeMap;

import valtoexcel.Constants;
import valtoexcel.ExcelPoi;
import valtoexcel.Val;

public class ValToExcelMain {

	public static void main(String[] args) throws IOException, SQLException {

		//		connection 
		String user ="NOVO";
		String pass = "novo";
		String url ="jdbc:oracle:thin:@//localhost:1521/xe";
			
		Val val1 = new Val( "VAL1",5 , 5, 4, Constants.QUERY1);
		Val val2 = new Val( "VAL2", 5, 5, 11, Constants.QUERY2);

		HashSet<Val> set = new HashSet<>();
		set.add(val1);
		set.add(val2);

		for (Val val : set) {

			SortedMap<Integer, List<String>> map = new TreeMap<>();
			try 
			(Connection c = DriverManager.getConnection(url, user, pass);
					Statement st = c.createStatement();
					)
			{				c.setReadOnly(true);
				System.out.println("Connected"  );

				try (
						ResultSet result =  st.executeQuery(val.getQuery());

						){

					int line = 0;
					while (result.next()) {
						line++;
						List<String> list2 = new ArrayList<>();
						for (int i = 1; i <= val.getMaxCollumn(); i++) {
						list2.add(result.getString(i)) ;

						}
						map.put(line,list2);
					}

				}
				val.setMap(map);
				System.out.println(val.getName() + " " +val.getMap().values());
			} catch (Exception e) {
				// TODO: handle exception
			}
			System.out.println(val2.getMap());

		}

		ExcelPoi.whiteMapValExel(set);

		//ExcelTutorial.whriteMapInExcel(map, "SQl");

	}

}


