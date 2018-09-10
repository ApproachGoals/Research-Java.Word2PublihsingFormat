package db;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Style;

public class AcademicStyleMap {

	Map<String, Style> styleMap;
	
	Map<String, AbstractNum> numberingMap;
	
	public Map<String, Style> getStyleMap() {
		return styleMap;
	}

	public void setStyleMap(Map<String, Style> styleMap) {
		this.styleMap = styleMap;
	}

	public Map<String, AbstractNum> getNumberingMap() {
		return numberingMap;
	}

	public void setNumberingMap(Map<String, AbstractNum> numberingMap) {
		this.numberingMap = numberingMap;
	}

	public AcademicStyleMap() {
		// TODO Auto-generated constructor stub
		styleMap = new HashMap<String, Style>();
		
		numberingMap = new HashMap<String, AbstractNum>();
		
	}
}
