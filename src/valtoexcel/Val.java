package valtoexcel;

import java.io.File;
import java.util.List;
import java.util.SortedMap;

public class Val {

	

	public SortedMap<Integer, List<String>> getMap() {
		return map;
	}

	public void setMap(SortedMap<Integer, List<String>> map) {
		this.map = map;
	}
	private String fullName;
	private StatusVal status;

	private String name;
	private int line = 1; 
	private int col = 1 ;
	private int maxCollumn;
	private SortedMap<Integer, List<String>> map;
	private List<String> headNames;
	private String query;


	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}


	public int getMaxCollumn() {
		return maxCollumn;
	}

	public void setMaxCollumn(int maxCollumn) {
		this.maxCollumn = maxCollumn;
	}

	public String getQuery() {
		return query;
	}

	public void setQuery(String query) {
		this.query = query;
	}
	
	public Val( String name, int line, int col , String query, String fullName) {
		super();
		//this.file = file;
		this.name = name;
		this.col= col;
		this.setLine(line);
	
		this.query = query;
		this.fullName = fullName;
	}
	public Val( String name, int line, int col) {
		super();
		//this.file = file;
		this.name = name;
		this.col= col;
		this.setLine(line);
	
		
	}
	public int getLine() {
		return line;
	}

	public void setLine(int line) {
		this.line = line;
	}

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}
	public StatusVal getStatus() {
		return status;
	}

	public void setStatus(StatusVal status) {
		this.status = status;
	}
	public String getFullName() {
		return fullName;
	}

	public void setFullName(String fullName) {
		this.fullName = fullName;
	}
	public List<String> getHeadNames() {
		return headNames;
	}

	public void setHeadNames(List<String> headNames) {
		this.headNames = headNames;
	}
	public enum StatusVal {
		
		OK ("OK"),
		NOK ("NOK");
		String okNok;
		private StatusVal(String okNok) {
			this.setOkNok(okNok);
		}
		String getOkNok() {
			return okNok;
		}
		void setOkNok(String okNok) {
			this.okNok = okNok;
		}
	
	}
}
