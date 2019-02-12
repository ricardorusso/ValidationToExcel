package valtoexcel;

public abstract class Constants {
	public static final String QUERY1 = "select /*+ PARALLEL(16)*/" + 
			"    *" + 
			"FROM" + 
			"    jobs";
	public static final String QUERY2 = "SELECT * FROM  employees";
	private Constants() {}
	
	
}
