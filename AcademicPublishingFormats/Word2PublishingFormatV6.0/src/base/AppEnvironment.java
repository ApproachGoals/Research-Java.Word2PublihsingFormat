package base;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Style;

import db.AcademicStylePool;
import ui.MainFrame;

public class AppEnvironment {

	public static File currentFile;
	
	private File newStyleFile;
	private Map<String, Style> newStyleMap;
	
	private MainFrame gui;
	
	private static AppEnvironment me;
	
	private boolean isGerman;
	
	private boolean isNeedUpdating;
	
	private boolean isActiveTestMode;
	
	private AcademicStylePool stylePool;
	
	private NumberingDefinitionsPart numberingPart;
	
	public AppEnvironment() {
		
	}
	
	public static AppEnvironment getInstance(){
		if(me==null){
			me = new AppEnvironment();
		}
		return me;
	}
	
	public void init() {
		currentFile = null;
	}
	
	public static File getCurrentFile() {
		return currentFile;
	}

	public static void setCurrentFile(File currentFile) {
		AppEnvironment.currentFile = currentFile;
	}

	public MainFrame getGui() {
		return gui;
	}

	public void setGui(MainFrame gui) {
		this.gui = gui;
	}

	public boolean isGerman() {
		return isGerman;
	}

	public void setGerman(boolean isGerman) {
		this.isGerman = isGerman;
	}

	public boolean isNeedUpdating() {
		return isNeedUpdating;
	}

	public void setNeedUpdating(boolean isNeedUpdating) {
		this.isNeedUpdating = isNeedUpdating;
	}

	public AcademicStylePool getStylePool() {
		return stylePool;
	}

	public void setStylePool(AcademicStylePool stylePool) {
		this.stylePool = stylePool;
	}

	public File getNewStyleFile() {
		return newStyleFile;
	}

	public void setNewStyleFile(File newStyleFile) {
		this.newStyleFile = newStyleFile;
	}

	public Map<String, Style> getNewStyleMap() {
		return newStyleMap;
	}

	public void setNewStyleMap(Map<String, Style> newStyleMap) {
		this.newStyleMap = newStyleMap;
	}

	public boolean isActiveTestMode() {
		return isActiveTestMode;
	}

	public void setActiveTestMode(boolean isActiveTestMode) {
		this.isActiveTestMode = isActiveTestMode;
	}

	public NumberingDefinitionsPart getNumberingPart() {
		return numberingPart;
	}

	public void setNumberingPart(NumberingDefinitionsPart numberingPart) {
		this.numberingPart = numberingPart;
	}

}
