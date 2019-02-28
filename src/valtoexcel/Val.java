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
	private Status status;
	private File file;
	private String name;
	private int line = 1; 
	private int col = 1 ;
	private int maxCollumn;
	private SortedMap<Integer, List<String>> map;

	private String query;

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}

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
	
	public Val( String name, int line, int col ,  int maxCollumn, String query) {
		super();
		//this.file = file;
		this.name = name;
		this.col= col;
		this.setLine(line);
		this.maxCollumn = maxCollumn;
		this.query = query;
	}
	public Val( String name, int line, int col ,  int maxCollumn) {
		super();
		//this.file = file;
		this.name = name;
		this.col= col;
		this.setLine(line);
		this.maxCollumn = maxCollumn;
		
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
	public Status getStatus() {
		return status;
	}

	public void setStatus(Status status) {
		this.status = status;
	}
	private enum Status {
		OK,
		NOK
	}
}
