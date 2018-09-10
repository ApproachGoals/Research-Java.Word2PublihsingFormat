package lib.ooxml.ref;

import java.math.BigInteger;
import java.util.ArrayList;

import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.CTVerticalAlignRun;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.NumFmt;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STHint;
import org.docx4j.wml.STTabJc;
import org.docx4j.wml.STVerticalAlignRun;
import org.docx4j.wml.Tabs;
import org.docx4j.wml.Lvl.Legacy;
import org.docx4j.wml.Lvl.LvlText;
import org.docx4j.wml.Lvl.PStyle;
import org.docx4j.wml.Lvl.Start;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.PPrBase.Ind;

public class OOXMLNumbering {

	public static AbstractNum createLvl(AbstractNum abstractNum, String lvlILvl, 
			NumberFormat numberFormat, String pStyle, String lvlText, 
			boolean isLegacy, String legacyValue, String legacyIndent,
			JcEnumeration lvlJc,
			STTabJc stTabJc, String tabsVal, 
			String indFirstLine, String indLeftVal, String hanging, String rightVal,
			String rFontVal, STHint stHint, 
			boolean isBold, boolean isItalic, boolean capsVal, boolean strikeVal, boolean dStrikeVal, boolean vanishVal,
			String colorVal, String sz,  String szcs, STVerticalAlignRun vAlign){
		
		Lvl lvl = getLvl(abstractNum, lvlILvl);
		boolean isDoAppend = false;
		if(lvl==null){
			lvl = new Lvl();
			isDoAppend = true;
		}
		lvl.setIlvl(new BigInteger(lvlILvl));
		
		setStart(lvl);
		setNumFmt(lvl, numberFormat);
		setPStyle(lvl, pStyle);
		setLvlText(lvl, lvlText);
		if(isLegacy){
			setLegacy(lvl, isLegacy, legacyValue, legacyIndent);
		}
		setLvlJc(lvl, lvlJc);
		
		lvl.setPPr(new PPr());
		if(stTabJc!=null){
			setTabs(lvl, stTabJc, tabsVal);
		}
		
		if(indFirstLine!=null || indLeftVal!=null || hanging!=null || rightVal!=null){
			setInd(lvl, indFirstLine, indLeftVal, hanging, rightVal);
		}
		
		lvl.setRPr(new RPr());
		if(rFontVal!=null){
			setRFonts(lvl, rFontVal, stHint);
		}
		
		setBoldAndItalic(lvl, isBold, isItalic);
		setCaps(lvl, true, capsVal);
		setStrike(lvl, true, strikeVal);
		setDstrike(lvl, true, dStrikeVal);
		setVanish(lvl, true, vanishVal);
		
		setColor(lvl, colorVal);
		
		setSZ(lvl, sz);
		setSzCs(lvl, szcs);
		
		setVertAlign(lvl, vAlign);
		
//		lvl.getRPr().setShadow(new BooleanDefaultTrue());
//		not namespace of w14
		if(isDoAppend){
			abstractNum.getLvl().add(lvl);
		}
		
		return abstractNum;
	}
	
	private static Lvl getLvl(AbstractNum abstractNum, String lvlILvl){
		Lvl lvl = null;
		for(int index=0; index<abstractNum.getLvl().size(); index++){
			if(abstractNum.getLvl().get(index).getIlvl().compareTo(new BigInteger(lvlILvl))==0){
				lvl = abstractNum.getLvl().get(index);
			}
		}
		return lvl;
	}
	
	public static void setLegacy(Lvl lvl, boolean isLegacy, String legacyValue, String legacyIndent){
		lvl.setLegacy(new Legacy());
		lvl.getLegacy().setLegacy(new Boolean(isLegacy));
		lvl.getLegacy().setLegacySpace(new BigInteger(legacyValue));
		lvl.getLegacy().setLegacyIndent(new BigInteger(legacyIndent));
	}
	
	public static void setBoldAndItalic(Lvl lvl, boolean isBold, boolean isItalic){
		if(!isBold){
			lvl.getRPr().setB(null);
			lvl.getRPr().setBCs(null);
		} else {
			lvl.getRPr().setB(new BooleanDefaultTrue());
			lvl.getRPr().setBCs(new BooleanDefaultTrue());
		}
		
		if(!isItalic){
			lvl.getRPr().setI(null);
			lvl.getRPr().setICs(null);
		} else {
			lvl.getRPr().setI(new BooleanDefaultTrue());
			lvl.getRPr().setICs(new BooleanDefaultTrue());
		}
	}
	
	public static void setVertAlign(Lvl lvl, STVerticalAlignRun val){
		lvl.getRPr().setVertAlign(new CTVerticalAlignRun());
		lvl.getRPr().getVertAlign().setVal(val);
	}
	
	public static void setSzCs(Lvl lvl, String szCsVal){
		lvl.getRPr().setSzCs(new HpsMeasure());
		lvl.getRPr().getSzCs().setVal(new BigInteger(szCsVal));
	}
	
	public static void setSZ(Lvl lvl, String szVal){
		lvl.getRPr().setSz(new HpsMeasure());
		lvl.getRPr().getSz().setVal(new BigInteger(szVal));
	}
	
	public static void setColor(Lvl lvl, String val){
		lvl.getRPr().setColor(new Color());
		lvl.getRPr().getColor().setVal(val);
	}
	
	public static void setVanish(Lvl lvl, boolean isVanish, boolean val){
		if(isVanish){
			lvl.getRPr().setVanish(new BooleanDefaultTrue());
			lvl.getRPr().getVanish().setVal(new Boolean(false));
		} else {
			lvl.getRPr().setVanish(null);
		}
	}
	
	public static void setDstrike(Lvl lvl, boolean isDStrike, boolean val){
		if(isDStrike){
			lvl.getRPr().setDstrike(new BooleanDefaultTrue());
			lvl.getRPr().getDstrike().setVal(new Boolean(val));
		} else {
			lvl.getRPr().setDstrike(null);
		}
	}
	
	public static void setStrike(Lvl lvl, boolean isStrike, boolean val){
		if(isStrike){
			lvl.getRPr().setStrike(new BooleanDefaultTrue());
			lvl.getRPr().getStrike().setVal(new Boolean(val));
		} else {
			lvl.getRPr().setStrike(null);
		}
		
	}
	
	public static void setCaps(Lvl lvl, boolean isCaps, boolean val){
		if(isCaps){
			lvl.getRPr().setCaps(new BooleanDefaultTrue());
			lvl.getRPr().getCaps().setVal(new Boolean(val));
		} else {
			lvl.getRPr().setCaps(null);
		}
		
	}
	
	public static void setRFonts(Lvl lvl, String fontVal, STHint stHint){
		lvl.getRPr().setRFonts(new RFonts());
		lvl.getRPr().getRFonts().setAscii(fontVal);
		lvl.getRPr().getRFonts().setHAnsi(fontVal);
		lvl.getRPr().getRFonts().setCs(fontVal);
		lvl.getRPr().getRFonts().setHint(stHint);
	}
	
	public static void setInd(Lvl lvl, String firstLineVal, String leftVal, String hangingVal, String rightVal){
		lvl.getPPr().setInd(new Ind());
		if(firstLineVal!=null){
			lvl.getPPr().getInd().setFirstLine(new BigInteger(firstLineVal));
		}
		if(leftVal!=null){
			lvl.getPPr().getInd().setLeft(new BigInteger(leftVal));
		}
		if(hangingVal!=null){
			lvl.getPPr().getInd().setHanging(new BigInteger(hangingVal));
		}
		if(rightVal!=null){
			lvl.getPPr().getInd().setRight(new BigInteger(rightVal));
		}
		
	}
	
	public static void setTabs(Lvl lvl, STTabJc type, String val){
		lvl.getPPr().setTabs(new Tabs());
		CTTabStop tab = new CTTabStop();
		tab.setVal(type);
		tab.setPos(new BigInteger(val));
		lvl.getPPr().getTabs().getTab().add(tab);
	}
	
	public static void setLvlJc(Lvl lvl, JcEnumeration lvlJc){
		lvl.setLvlJc(new Jc());
		lvl.getLvlJc().setVal(lvlJc);
	}
	
	public static void setLvlText(Lvl lvl, String lvlText){
		lvl.setLvlText(new LvlText());
		lvl.getLvlText().setVal(lvlText);
	}
	
	public static void setPStyle(Lvl lvl, String pStyle){
		lvl.setPStyle(new PStyle());
		lvl.getPStyle().setVal(pStyle);
	}
	
	public static void setStart(Lvl lvl){
		lvl.setStart(new Start());
		lvl.getStart().setVal(new BigInteger("1"));
	}
	
	public static void setNumFmt(Lvl lvl, NumberFormat numberFormat){
		lvl.setNumFmt(new NumFmt());
		lvl.getNumFmt().setVal(numberFormat);
	}
	
}
