package squash.model;

import java.util.ArrayList;
import java.util.List;

public class SquashReportData {
	private List<LineReportSquash> linesOfcontent;
	
	public SquashReportData() {
		linesOfcontent = new ArrayList<LineReportSquash>();
	}
	
	public void addLineOfContent(List<String> content){
		linesOfcontent.add(new LineReportSquash(content.get(0), content.get(1), content.get(2), content.get(3), content.get(4), content.get(5), content.get(6), content.get(7), content.get(8), content.get(9), content.get(10), content.get(11), content.get(12), content.get(13), content.get(14), content.get(15)));
	}
	
	public List<LineReportSquash> getLinesOfContent() {
		return linesOfcontent;
	}
}
