package lib.ooxml.style;

import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.Style;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.Style.BasedOn;
import org.docx4j.wml.Style.Name;
import org.docx4j.wml.Style.Next;

import base.AppEnvironment;
import base.StaticConstants;
import db.AcademicFormatStructureDefinition;
import db.AcademicStylePool;
import lib.AbstractPublishingStyle;
import lib.ooxml.OOXMLConstantString;

public class OOXMLSpringerStyle implements AbstractPublishingStyle {
	
	private ArrayList<Style> styleList;
	private Map<String, Style> styleMap; 
	
	public ArrayList<Style> getStyleList() {
		return styleList;
	}

	public void setStyleList(ArrayList<Style> styleList) {
		this.styleList = styleList;
	}
	
	public Map<String, Style> getStyleMap() {
		return styleMap;
	}

	public void setStyleMap(Map<String, Style> styleMap) {
		this.styleMap = styleMap;
	}
	public OOXMLSpringerStyle() {
		// TODO Auto-generated constructor stub
		styleList = new ArrayList<Style>();
		styleMap = new HashMap<String, Style>();
		Style style = new Style();
		/*
		 * Normal
		 * 10-point
		 * Times New Roman
		 */
		style = new Style();
		style.setStyleId(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setName(new Name());
		style.getName().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times");	// Times Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_DEFAULTPARAGRAPH, style);
		

		/*
		 * Title
		 * Times New Roman
		 * bold
		 * 14 points
		 */
		style = new Style();
		style.setStyleId("title");
		style.setName(new Name());
		style.getName().setVal("title");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setNext(new Next());
		style.getNext().setVal("author");
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("460"));
		style.getPPr().getSpacing().setLine(new BigInteger("348"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setPageBreakBefore(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("28"));	// 14-points
		style.getRPr().setB(new BooleanDefaultTrue());		// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TITLE, style);
		
		/*
		 * Authors		
		 * basedOn Standard
		 */
		style = new Style();
		style.setStyleId("author");
		style.setName(new Name());
		style.getName().setVal("author");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setNext(new Next());
		style.getNext().setVal("authorinfo");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));// 10-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AUTHORS, style);
		
		/*
		 * authorinfo
		 * 9-point
		 */
		style = new Style();
		style.setStyleId("authorinfo");
		style.setName(new Name());
		style.getName().setVal("authorinfo");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setNext(new Next());
		style.getNext().setVal("email");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(BigInteger.valueOf(0));
		style.getPPr().getSpacing().setBefore(BigInteger.valueOf(0));
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AFFILIATIONS, style);
		
		/*
		 * email
		 * 9-point
		 */
		style = new Style();
		style.setStyleId("email");
		style.setName(new Name());
		style.getName().setVal("email");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setNext(new Next());
		style.getNext().setVal("abstract");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_EMAIL, style);
		
		/*
		 * abstract
		 * 
		 */
		style = new Style();
		
		style.setStyleId("abstract");
		style.setName(new Name());
		style.getName().setVal("abstract");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
//		style.setNext(new Next());
//		style.getNext().setVal("heading1");
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("600"));
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		style.getPPr().setInd(new Ind());
		style.getPPr().getInd().setLeft(new BigInteger("567"));
		style.getPPr().getInd().setRight(new BigInteger("567"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_ABSTRACT, style);
		
		/*
		 * keyword head
		 * 9-point
		 * bold
		 * center
		 */
		style = new Style();
		style.setStyleId("Keywordhead");
		style.setName(new Name());
		style.getName().setVal("Keywordhead");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("600"));
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		style.getPPr().setInd(new Ind());
		style.getPPr().getInd().setLeft(new BigInteger("567"));
		style.getPPr().getInd().setRight(new BigInteger("567"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_KEYWORDHEAD, style);
		
		/*
		 * heading1
		 * basedOn w:val="Standard"
		 * keepNext
		 * keepLines
		 * suppressAutoHyphens
		 * spacing w:before="520" w:after="280"
		 * bold
		 * sz w:val="24"
		 */
		style = new Style();
		
		style.setStyleId("heading1");
		style.setName(new Name());
		style.getName().setVal("heading1");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("520"));
		style.getPPr().getSpacing().setAfter(new BigInteger("280"));
		style.getPPr().setInd(new Ind());
		style.getPPr().getInd().setLeft(new BigInteger("0"));
//		style.getPPr().getInd().setRight(new BigInteger("567"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("0"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING1, style);
		
		/*
		 * heading2
		 * basedOn w:val="Standard"
		 * keepNext
		 * keepLines
		 * suppressAutoHyphens
		 * spacing w:before="440" w:after="220"
		 */
		style = new Style();
		
		style.setStyleId("heading2");
		style.setName(new Name());
		style.getName().setVal("heading2");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("440"));
		style.getPPr().getSpacing().setAfter(new BigInteger("220"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("1"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING2, style);
		
		/*
		 * heading3
		 * basedOn w:val="Standard"
		 * keepNext
		 * keepLines
		 * suppressAutoHyphens
		 * spacing w:before="320"
		 */
		style = new Style();
		style.setStyleId("heading3");
		style.setName(new Name());
		style.getName().setVal("heading3");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(Context.getWmlObjectFactory().createStyleBasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING3, style);
		
		style = new Style();
		
		style.setStyleId("heading3Char");
		style.setName(new Name());
		style.getName().setVal("heading3Char");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("320"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));// 10-points
		style.getRPr().setB(new BooleanDefaultTrue());
		
		style.getPPr().setNumPr(null);
//		style.getPPr().getNumPr().setNumId(new NumId());
//		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
//		style.getPPr().getNumPr().setIlvl(new Ilvl());
//		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("2"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING3CHAR, style);

		/*
		 * heading4
		 * italic 10px
		 */
		style = new Style();
		style.setStyleId("heading4");
		style.setName(new Name());
		style.getName().setVal("heading4");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(Context.getWmlObjectFactory().createStyleBasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING4, style);
		
		style = new Style();
		
		style.setStyleId("heading4Char");
		style.setName(new Name());
		style.getName().setVal("heading4Char");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("320"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());
		
		style.getPPr().setNumPr(null);
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING4CHAR, style);
		
		/*
		 * heading5
		 * basedOn w:val="Standard"
		 * keepNext
		 * keepLines
		 * suppressAutoHyphens
		 * spacing w:before="520" w:after="280"
		 * bold
		 * sz w:val="24"
		 */
		style = new Style();
		
		style.setStyleId("heading5");
		style.setName(new Name());
		style.getName().setVal("heading5");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("520"));
		style.getPPr().getSpacing().setAfter(new BigInteger("280"));
		style.getPPr().setInd(new Ind());
		style.getPPr().getInd().setLeft(new BigInteger("0"));
//		style.getPPr().getInd().setRight(new BigInteger("567"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());
		
		style.getPPr().setNumPr(null);
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING5, style);
		
		/*
		 * reference
		 * ind w:left="227" w:hanging="227"
		 * sz w:val="18"
		 */
		style = new Style();
		style.setStyleId("reference");
		style.setName(new Name());
		style.getName().setVal("reference");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
//		style.setBasedOn(Context.getWmlObjectFactory().createStyleBasedOn());
//		style.getBasedOn().setVal("heading1");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setInd(new Ind());
		style.getPPr().getInd().setLeft(new BigInteger("227"));
		style.getPPr().getInd().setHanging(new BigInteger("227"));
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));
		style.getRPr().setB(new BooleanDefaultTrue());
		
//		style.getPPr().setNumPr(null);
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCEHEAD, style);
		
		/*
		 * referenceitem
		 * ind w:left="227" w:hanging="227"
		 * sz w:val="18"
		 */
		style = new Style();
		style.setStyleId("referenceitem");
		style.setName(new Name());
		style.getName().setVal("referenceitem");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setInd(new Ind());
		style.getPPr().getInd().setLeft(new BigInteger("227"));
		style.getPPr().getInd().setHanging(new BigInteger("227"));
		
		style.setRPr(new RPr());
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("0"));
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_REFERENCE_NUMBERING_ABSTRACTNUMID));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCE, style);

		/*
		 * tabletitle
		 * keepNext
		 * keepLines
		 * spacing w:before="240" w:after="120"
		 * sz w:val="18"
		 */
		style = new Style();
		style.setStyleId("tabletitle");
		style.setName(new Name());
		style.getName().setVal("tabletitle");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("240"));
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		
		style.setRPr(new RPr());
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TABLEHEADER, style);
		
		style = new Style();
		style.setStyleId("tabletitleChar");
		style.setName(new Name());
		style.getName().setVal("tabletitleChar");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		
		style.setRPr(new RPr());
		style.getRPr().setB(new BooleanDefaultTrue());
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TABLEHEADERCHAR, style);
		
		/*
		 * caption
		 * next w:val="Standard"
		 * qFormat
		 * spacing w:before="120" w:after="120"
		 * b
		 */
		style = new Style();
		style.setStyleId("caption");
		style.setName(new Name());
		style.getName().setVal("caption");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("120"));	
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		
		style.setRPr(new RPr());
//		style.getRPr().setSz(new HpsMeasure());
//		style.getRPr().getSz().setVal(new BigInteger("18"));
		style.getRPr().setB(new BooleanDefaultTrue());
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_CAPTION, style);
		
		/*
		 * Figure
		 * Times New Roman 
		 * 9-point 
		 * bold		
		 */
		style = new Style();
		style.setStyleId("Figure");
		style.setName(new Name());
		style.getName().setVal("Figure");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("120"));	
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		
		style.setRPr(new RPr());
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));
//		style.getRPr().setB(new BooleanDefaultTrue());
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FIGURE, style);
		
		style = new Style();
		style.setStyleId("FigureChar");
		style.setName(new Name());
		style.getName().setVal("FigureChar");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setRPr(new RPr());
		style.getRPr().setB(new BooleanDefaultTrue());
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FIGURECHAR, style);
		/*
		 * hyperlink
		 */
		style = new Style();
		style.setStyleId("Hyperlink");
		style.setName(new Name());
		style.getName().setVal("Hyperlink");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);		// flush both
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("0000FF");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times");	// Times New Roman
		style.getRPr().setNoProof(new BooleanDefaultTrue());
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HYPERLINK, style);
	}

	@Override
	public void loadStyleFromFile(File file) {
		// TODO Auto-generated method stub
		if(file==null){
			if(AppEnvironment.getInstance().getStylePool()!=null
					&& AppEnvironment.getInstance().getStylePool().getStyles()!=null
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER)!=null
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap()!=null){
				styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap();
			} else {
				
			}
		} else {
			AcademicStylePool stylePool = new AcademicStylePool();
			stylePool.loadFile(file);
			styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap();
		}
	}

}
