package lib.ooxml.ref;

import java.math.BigInteger;
import java.util.List;
import java.util.Map;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.CTLongHexNumber;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.STHint;
import org.docx4j.wml.STTabJc;
import org.docx4j.wml.STVerticalAlignRun;
import org.docx4j.wml.Style;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.Num;
import org.docx4j.wml.Numbering.AbstractNum.MultiLevelType;
import org.docx4j.wml.Numbering.Num.AbstractNumId;

import base.AppEnvironment;
import base.StaticConstants;
import db.AcademicFormatStructureDefinition;

public class IEEENumbering extends OOXMLNumbering {
	public static void generateCaptionNumbering(NumberingDefinitionsPart numberingDefinitionsPart, boolean isTable){
		try {
			Numbering numbering = numberingDefinitionsPart.getContents();
			List<AbstractNum> abNumList = numbering.getAbstractNum();
			List<Num> numList = numbering.getNum();
			ObjectFactory wmlObjectFactory = new ObjectFactory();
			boolean isDoAppend = false;
			
			AbstractNum abstractNumTemp = isTable?
					AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGCAPTIONSTRING)
					:AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGFIGURESTRING);
			Map<String, Style> stylemap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getStyleMap();
			
			if(abstractNumTemp!=null){
				
			} else {
				abstractNumTemp = wmlObjectFactory.createNumberingAbstractNum();
			}
			if(abstractNumTemp.getAbstractNumId()!=null){
				
			} else {
				if(isTable){
					abstractNumTemp.setAbstractNumId(new BigInteger(StaticConstants.OOXML_TABLE_HEADER_NUMBERING_ABSTRACTNUMID));
				} else {
					abstractNumTemp.setAbstractNumId(new BigInteger(StaticConstants.OOXML_FIGURE_HEADER_NUMBERING_ABSTRACTNUMID));
				}
			}
			AbstractNum abstractNum = getAbstractNum(abNumList, abstractNumTemp.getAbstractNumId().toString());
			
			if(abstractNum==null){
				abstractNum = wmlObjectFactory.createNumberingAbstractNum();
				abstractNum.setAbstractNumId(abstractNumTemp.getAbstractNumId());
				isDoAppend = true;
			}
			
			MultiLevelType lvlType = new MultiLevelType();
			lvlType.setVal("singleLevel");
			abstractNum.setMultiLevelType(lvlType);
			CTLongHexNumber longhexnumber = wmlObjectFactory.createCTLongHexNumber();
			abstractNum.setNsid(longhexnumber);
			longhexnumber.setVal( isTable?"6ECB300C":"6ECB300D" );
			
			List<Lvl> lvlList = abstractNumTemp.getLvl();
			for(Lvl lvl : lvlList){
				if(lvl.getIlvl()!=null){
					String name = isTable?AcademicFormatStructureDefinition.CAPTIONHEADER:AcademicFormatStructureDefinition.FIGUREHEADER;
					
					createLvl(abstractNum, "0", lvl.getNumFmt().getVal(), name, 
							lvl.getLvlText().getVal(), false, null, null, JcEnumeration.LEFT, 
							STTabJc.NUM, "0", 
							(stylemap.get(name).getPPr().getInd()!=null && stylemap.get(name).getPPr().getInd().getLeft()!=null)?(stylemap.get(name).getPPr().getInd().getLeft().intValue()+""):"0", 
							(stylemap.get(name).getPPr().getInd()!=null && stylemap.get(name).getPPr().getInd().getHanging()!=null)?(stylemap.get(name).getPPr().getInd().getHanging().intValue()+""):null, 
							null, null, 
							stylemap.get(name).getRPr().getRFonts()!=null?stylemap.get(name).getRPr().getRFonts().getAscii():null, 
							STHint.DEFAULT, 
							stylemap.get(name).getRPr().getB()!=null?stylemap.get(name).getRPr().getB().isVal():false, 
							stylemap.get(name).getRPr().getI()!=null?stylemap.get(name).getRPr().getI().isVal():false,
							stylemap.get(name).getRPr().getCaps()!=null?stylemap.get(name).getRPr().getCaps().isVal():false,
							false, false, false,
							"auto", 
							stylemap.get(name).getRPr().getSz()!=null?stylemap.get(name).getRPr().getSz().getVal().toString():"20", 
							stylemap.get(name).getRPr().getSz()!=null?stylemap.get(name).getRPr().getSz().getVal().toString():"20", 
							STVerticalAlignRun.BASELINE);
				}
			}
			
			if(isDoAppend){
				Num appendNum = new Num();
				appendNum.setNumId(abstractNumTemp.getAbstractNumId());
				AbstractNumId numId = new AbstractNumId();
				numId.setVal(abstractNum.getAbstractNumId());
				appendNum.setAbstractNumId(numId);
				numList.add(appendNum);
				abNumList.add(abstractNum);
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
	public static void generateReferenceNumbering(NumberingDefinitionsPart numberingDefinitionsPart){
		try {
			if(numberingDefinitionsPart!=null){
				
			} else {
				return;
			}
			Numbering numbering = numberingDefinitionsPart.getContents();
			List<AbstractNum> abNumList = numbering.getAbstractNum();
			List<Num> numList = numbering.getNum();
			ObjectFactory wmlObjectFactory = new ObjectFactory();
			AbstractNum abstractNum = getAbstractNum(abNumList, StaticConstants.OOXML_REFERENCE_NUMBERING_ABSTRACTNUMID);
			boolean isDoAppend = false;
			if(abstractNum==null){
				abstractNum = wmlObjectFactory.createNumberingAbstractNum();
				abstractNum.setAbstractNumId(new BigInteger(StaticConstants.OOXML_REFERENCE_NUMBERING_ABSTRACTNUMID));
				isDoAppend = true;
			}
			
			MultiLevelType lvlType = new MultiLevelType();
			lvlType.setVal("singleLevel");
			abstractNum.setMultiLevelType(lvlType);
			CTLongHexNumber longhexnumber = wmlObjectFactory.createCTLongHexNumber();
			abstractNum.setNsid(longhexnumber);
			longhexnumber.setVal( "6ECB300B");
			
			String styleName;
			
			if(AppEnvironment.getInstance().isGerman()){
				styleName = "Literaturverzeichnis";
			} else {
				styleName = "References";
			}
			/*
			 * createLvl(
			 * abstractNum, 
			 * lvlILvl, 
			 * numberFormat, 		
			 * pStyle, lvlText, 
			 * isLegacy, legacyValue, legacyIndent
			 * lvlJc, 			
			 * stTabJc, 	tabsVal, 
			 * indFirstLine, indLeftVal, hanging, rightVal, 
			 * rFontVal, stHint, 
			 * capsVal, strikeVal, dStrikeVal, vanishVal, 
			 * colorVal, 
			 * sz, szcs, 
			 * vAlign)
			 */
			createLvl(abstractNum, "0", NumberFormat.DECIMAL, styleName, "[%1]", false, "360", "0", JcEnumeration.LEFT, 
					STTabJc.NUM, "360", null, "0", "360", null, "Times New Roman", STHint.DEFAULT, 
					false, false, false, false, false, false,
					"000000", "16", "16", STVerticalAlignRun.BASELINE);
			
			if(isDoAppend){
				Num appendNum = new Num();
				appendNum.setNumId(abstractNum.getAbstractNumId());
				AbstractNumId numId = new AbstractNumId();
				numId.setVal(abstractNum.getAbstractNumId());
				appendNum.setAbstractNumId(numId);
				numList.add(appendNum);
				abNumList.add(abstractNum);
			}
		
			
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
	
	private static void generateSectionHeaderNumbering(NumberingDefinitionsPart numberingDefinitionsPart){
		try {
			Numbering numbering = numberingDefinitionsPart.getContents();
			
			AbstractNum abstractNumTemp = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGSECTIONSTRING);
			Map<String, Style> stylemap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getStyleMap();
			
			List<AbstractNum> abNumList = numbering.getAbstractNum();
			List<Num> numList = numbering.getNum();
			ObjectFactory wmlObjectFactory = new ObjectFactory();
			
			if(abstractNumTemp.getAbstractNumId()!=null){
				
			} else {
				abstractNumTemp.setAbstractNumId(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
			}
			AbstractNum abstractNum = getAbstractNum(abNumList, abstractNumTemp.getAbstractNumId().toString());
			boolean isDoAppend = false;
			if(abstractNum==null){
				abstractNum = wmlObjectFactory.createNumberingAbstractNum();
//				abstractNum.setAbstractNumId(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
				abstractNum.setAbstractNumId(abstractNumTemp.getAbstractNumId());
				isDoAppend = true;
			}
			
			MultiLevelType lvlType = new MultiLevelType();
			lvlType.setVal("multilevel");
			abstractNum.setMultiLevelType(lvlType);
			CTLongHexNumber longhexnumber = wmlObjectFactory.createCTLongHexNumber();
			abstractNum.setNsid(longhexnumber);
			longhexnumber.setVal( "6ECB300A");
			
			List<Lvl> lvlList = abstractNumTemp.getLvl();
			for(Lvl lvl : lvlList){
				if(lvl.getIlvl()!=null){
					switch (lvl.getIlvl().intValue()) {
					case 0:
						createLvl(abstractNum, "0", lvl.getNumFmt().getVal(), AcademicFormatStructureDefinition.HEADING1, 
								lvl.getLvlText().getVal(), false, null, null, JcEnumeration.LEFT, 
								STTabJc.NUM, "567", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING1).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING1).getPPr().getInd().getLeft()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING1).getPPr().getInd().getLeft().intValue()+""):"0", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING1).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING1).getPPr().getInd().getHanging()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING1).getPPr().getInd().getHanging().intValue()+""):null, 
								null, null, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getRFonts()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getRFonts().getAscii():null, 
								STHint.DEFAULT, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getB()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getB().isVal():false, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getI()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getI().isVal():false,
								stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getCaps()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getCaps().isVal():false,
								false, false, false,
								"auto", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getSz().getVal().toString():"20", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING1).getRPr().getSz().getVal().toString():"20", 
								STVerticalAlignRun.BASELINE);
						break;
					case 1:
						createLvl(abstractNum, "1", lvl.getNumFmt().getVal(), AcademicFormatStructureDefinition.HEADING2, 
								lvl.getLvlText().getVal(), false, null, null, JcEnumeration.LEFT, 
								STTabJc.NUM, "360", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING2).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING2).getPPr().getInd().getLeft()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING2).getPPr().getInd().getLeft().intValue()+""):"288", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING2).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING2).getPPr().getInd().getHanging()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING2).getPPr().getInd().getHanging().intValue()+""):"288", 
								null, null, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getRFonts()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getRFonts().getAscii():null, 
								STHint.DEFAULT, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getB()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getB().isVal():false, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getI()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getI().isVal():true,
								stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getCaps()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getCaps().isVal():false,
								false, false, false,
								"auto", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getSz().getVal().toString():"20", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING2).getRPr().getSz().getVal().toString():"20", 
								STVerticalAlignRun.BASELINE);
						break;
					case 2:
						createLvl(abstractNum, "2", lvl.getNumFmt().getVal(), AcademicFormatStructureDefinition.HEADING3CHAR, 
								lvl.getLvlText().getVal(), false, null, null, JcEnumeration.LEFT, 
								STTabJc.NUM, "540", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getPPr().getInd().getLeft()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING3).getPPr().getInd().getLeft().intValue()+""):"180", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getPPr().getInd().getHanging()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING3).getPPr().getInd().getHanging().intValue()+""):null, 
								null, null, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getRFonts()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getRFonts().getAscii():null, 
								STHint.DEFAULT, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getB()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getB().isVal():false, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getI()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getI().isVal():true,
								stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getCaps()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getCaps().isVal():false,
								false, false, false,
								"auto", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getSz().getVal().toString():"20", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING3CHAR).getRPr().getSz().getVal().toString():"20", 
								STVerticalAlignRun.BASELINE);
						break;
					case 3:
						createLvl(abstractNum, "3", lvl.getNumFmt().getVal(), AcademicFormatStructureDefinition.HEADING4, 
								lvl.getLvlText().getVal(), false, null, null, JcEnumeration.LEFT, 
								STTabJc.NUM, "720", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING4).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING4).getPPr().getInd().getLeft()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING4).getPPr().getInd().getLeft().intValue()+""):"360", 
								(stylemap.get(AcademicFormatStructureDefinition.HEADING4).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING4).getPPr().getInd().getHanging()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING4).getPPr().getInd().getHanging().intValue()+""):null, 
								null, null, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getRFonts()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getRFonts().getAscii():null, 
								STHint.DEFAULT, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getB()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getB().isVal():false, 
								stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getI()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getI().isVal():true,
								stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getCaps()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getCaps().isVal():false,
								false, false, false,
								"auto", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getSz().getVal().toString():"20", 
								stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING4).getRPr().getSz().getVal().toString():"20", 
								STVerticalAlignRun.BASELINE);
						break;
					case 4:
//						createLvl(abstractNum, "4", lvl.getNumFmt().getVal(), AcademicFormatStructureDefinition.HEADING5, 
//								lvl.getLvlText().getVal(), false, null, null, JcEnumeration.LEFT, 
//								STTabJc.NUM, "720", 
//								(stylemap.get(AcademicFormatStructureDefinition.HEADING5).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING5).getPPr().getInd().getLeft()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING5).getPPr().getInd().getLeft().intValue()+""):"360", 
//								(stylemap.get(AcademicFormatStructureDefinition.HEADING5).getPPr().getInd()!=null && stylemap.get(AcademicFormatStructureDefinition.HEADING5).getPPr().getInd().getHanging()!=null)?(stylemap.get(AcademicFormatStructureDefinition.HEADING5).getPPr().getInd().getHanging().intValue()+""):null, 
//								null, null, 
//								stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getRFonts()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getRFonts().getAscii():null, 
//								STHint.DEFAULT, 
//								stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getB()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getB().isVal():false, 
//								stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getI()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getI().isVal():true,
//								stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getCaps()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getCaps().isVal():false,
//								false, false, false,
//								"auto", 
//								stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getSz().getVal().toString():"20", 
//								stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getSz()!=null?stylemap.get(AcademicFormatStructureDefinition.HEADING5).getRPr().getSz().getVal().toString():"20", 
//								STVerticalAlignRun.BASELINE);
//						createLvl(abstractNum, "4", NumberFormat.NONE, head5, "", false, null, null, JcEnumeration.LEFT, 
//								STTabJc.NUM, "3240", null, "2880", null, null, "Times New Roman", STHint.DEFAULT, 
//								false, false, false, false, false, false,
//								"auto", "20", "20", STVerticalAlignRun.BASELINE);
						break;
					default:
						break;
					}
				}
			}
			
/*			
			createLvl(abstractNum, "0", NumberFormat.UPPER_ROMAN, head1, "%1.", false, null, null, JcEnumeration.CENTER, 
//					STTabJc.NUM, "576", "216", null, null, null, "Times New Roman", STHint.DEFAULT,
					null, "576", null, null, null, null, "Times New Roman", STHint.DEFAULT,
					false, false, false, false, false, false,
					"auto", "20", "20", STVerticalAlignRun.BASELINE);
			
			createLvl(abstractNum, "1", NumberFormat.UPPER_LETTER, head2, "%2.", false, null, null, JcEnumeration.LEFT, 
//					STTabJc.NUM, "360", null, "288", "288", null, "Times New Roman", STHint.DEFAULT, 
					null, "360", null, "288", "288", null, "Times New Roman", STHint.DEFAULT, 
					false, true, false, false, false, false,
					"auto", "20", "20", STVerticalAlignRun.BASELINE);
			
			createLvl(abstractNum, "2", NumberFormat.DECIMAL, head3, "%3)", false, null, null, JcEnumeration.LEFT, 
					STTabJc.NUM, "540", "180", null, null, null, "Times New Roman", STHint.DEFAULT, 
					false, true, false, false, false, false,
					"auto", "20", "20", STVerticalAlignRun.BASELINE);
			
			createLvl(abstractNum, "3", NumberFormat.LOWER_LETTER, head4, "%4)", false, null, null, JcEnumeration.LEFT, 
					STTabJc.NUM, "720", "360", null, null, null, "Times New Roman", STHint.DEFAULT, 
					false, true, false, false, false, false,
					"auto", "20", "20", STVerticalAlignRun.BASELINE);
			
			createLvl(abstractNum, "4", NumberFormat.NONE, head5, "", false, null, null, JcEnumeration.LEFT, 
					STTabJc.NUM, "3240", null, "2880", null, null, "Times New Roman", STHint.DEFAULT, 
					false, false, false, false, false, false,
					"auto", "20", "20", STVerticalAlignRun.BASELINE);
			*/
			if(isDoAppend){
				Num appendNum = new Num();
				appendNum.setNumId(abstractNum.getAbstractNumId());
				AbstractNumId numId = new AbstractNumId();
				numId.setVal(abstractNum.getAbstractNumId());
				appendNum.setAbstractNumId(numId);
				numList.add(appendNum);
				abNumList.add(abstractNum);
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}

	private static AbstractNum getAbstractNum(List<AbstractNum> abstractNumList, String numId) {
		AbstractNum abstractNum = null;
		for(int index=0; index<abstractNumList.size(); index++) {
			AbstractNum temp = abstractNumList.get(index);
//			System.out.println(temp.getAbstractNumId()+"  "+numId);
			if (temp.getAbstractNumId().intValue() == (new BigInteger(numId)).intValue()) {
				abstractNum = temp;
				break;
			}
		}
		return abstractNum;
	}
	
	public static void adaptNumberingFile(WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart){
		NumberingDefinitionsPart numberingDefinitionsPart = documentPart.getNumberingDefinitionsPart();
		
		try {
			if(numberingDefinitionsPart!=null){
				
			} else {
				numberingDefinitionsPart = new NumberingDefinitionsPart();
				wordMLPackage.addTargetPart(numberingDefinitionsPart);
				numberingDefinitionsPart.setContents(Context.getWmlObjectFactory().createNumbering());
			}
			generateReferenceNumbering(numberingDefinitionsPart);
			generateSectionHeaderNumbering(numberingDefinitionsPart);

			generateCaptionNumbering(numberingDefinitionsPart, true);
			generateCaptionNumbering(numberingDefinitionsPart, false);
			
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		
	}
	
//	public static void modifyNumbering(MainDocumentPart documentPart, OOXMLNum num){
//		NumberingDefinitionsPart numberingDefinitionsPart = documentPart.getNumberingDefinitionsPart();
//		
//		try {
//			if(num!=null){
//				Numbering numbering = numberingDefinitionsPart.getContents();
//				List<AbstractNum> abNumList = numbering.getAbstractNum();
//				for(int index = 0; index < abNumList.size(); index++) {
//					AbstractNum abstractNum = abNumList.get(index);
//					if(abstractNum.getAbstractNumId().equals(num.getNumId())){
//						List<Lvl> lvlList = abstractNum.getLvl();
//						for(Lvl lvl : lvlList){
//							if(num.getiLvl()!=null && lvl.getIlvl().equals(num.getiLvl())){
//								lvl.setLvlText(new LvlText());
//								lvl.getLvlText().setVal("[%"+(num.getiLvl().intValue()+1)+"]");
//								if(lvl.getPPr()!=null){
//									
//								} else {
//									lvl.setPPr(new PPr());
//								}
//								lvl.getPPr().setInd(new Ind());
//								lvl.getPPr().getInd().setLeft(new BigInteger("360"));
//								lvl.getPPr().getInd().setHanging(new BigInteger("360"));
//								
//								if(lvl.getRPr()!=null){
//									lvl.getRPr().setB(null);
//									lvl.getRPr().setI(null);
//								}
//							}
//						}
//					}
//				}
//			}
//			
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//	}
}
