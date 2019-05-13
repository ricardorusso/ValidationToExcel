package valtoexcel;

import java.util.ArrayList;
import java.util.List;

public class Resume {
	
	private String valName, headResults;
	private int totalResult ; 
	private List<String> listValue = new ArrayList<>();
	private StringBuilder resumeStrB;
	public String getValName() {
		return valName;
	}
	public void setValName(String valName) {
		this.valName = valName;
	}

	public void setTotalResult(int totalResult) {
		this.totalResult = totalResult;
	}
	public List<String> getListValue() {
		return listValue;
	}
	public void setListValue(List<String> listValue) {
		this.listValue = listValue;
	}
	public StringBuilder getResumeStrB() {
		return resumeStrB;
	}
	public void setResumeStrB(StringBuilder resumeStrB) {
		this.resumeStrB = resumeStrB;
	}
	@Override
	public String toString() {
		
		return getValName() +" " + getResumeStrB() + " ";
		
	}
	
	public String toStringList(List<String> list) {
		if (list.isEmpty()) {
			return "";
		}
		StringBuilder strFinal = new StringBuilder(list.get(0)+": ");
		for (int i = 1; i < list.size(); i++) {
			strFinal.append(list.get(i) + ( (i==list.size()-1) ? " " : ", "));   
		}
		
		return strFinal.toString();
	}

}
