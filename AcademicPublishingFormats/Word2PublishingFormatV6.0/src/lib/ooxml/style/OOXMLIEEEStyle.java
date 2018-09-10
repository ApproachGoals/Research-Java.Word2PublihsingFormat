package lib.ooxml.style;

import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STTabJc;
import org.docx4j.wml.Style;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.Style.BasedOn;
import org.docx4j.wml.Style.Name;
import org.docx4j.wml.Tabs;

import base.AppEnvironment;
import base.StaticConstants;
import db.AcademicFormatStructureDefinition;
import db.AcademicStylePool;
import lib.AbstractPublishingStyle;
import lib.ooxml.OOXMLConstantString;

public class OOXMLIEEEStyle implements AbstractPublishingStyle {

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

	public OOXMLIEEEStyle() {
		// TODO
		// 1. abstract header italic
		
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
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_DEFAULTPARAGRAPH, style);
		
		/*
		 * Title
		 * 24-point
		 * Times New Roman
		 * bold
		 */
		style = new Style();
		style.setStyleId("Title");
		style.setName(new Name());
		style.getName().setVal("Title");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("48"));	// 24-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TITLE, style);
		
		/*
		 * Subtitle
		 * 14-point
		 * Times New Roman
		 */
		style = new Style();
		style.setStyleId("Subtitle");
		style.setName(new Name());
		style.getName().setVal("Subtitle");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("28"));	// 14-points
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_SUBTITLE, style);
		
		/*
		 * Authors
		 * 11-point
		 * Times New Roman
		 * center
		 */
		style = new Style();
		style.setStyleId("Authors");
		style.setName(new Name());
		style.getName().setVal("Authors");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("22"));	// 11-points
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AUTHORS, style);
		
		/*
		 * Affiliation
		 * 10-point
		 * Times New Roman
		 */
		style = new Style();
		style.setStyleId("Affiliation");
		style.setName(new Name());
		style.getName().setVal("Affiliation");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AFFILIATIONS, style);
		
		/*
		 * Abstract
		 * 9-point
		 * bold
		 * center
		 */
		style = new Style();
		style.setStyleId("Abstract");
		style.setName(new Name());
		style.getName().setVal("Abstract");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
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
		 * Footnotes 
		 * Times New Roman 
		 * 8-point
		 * fullwidth of page
		 */
		style = new Style();
		style.setStyleId("footnote");
		style.setName(new Name());
		style.getName().setVal("footnote");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("0"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FOOTNOTES, style);
		
		/*
		 * References
		 * Times New Roman 
		 * 8-point
		 * fullwidth of column
		 */
		style = new Style();
		style.setStyleId("References");
		style.setName(new Name());
		style.getName().setVal("References");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("0"));
		style.getPPr().getSpacing().setBefore(new BigInteger("0"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("0"));
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_REFERENCE_NUMBERING_ABSTRACTNUMID));
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setU(null);
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCE, style);
		
		/*
		 * ReferenceHead
		 * 10-point
		 * smallCaps
		 * center
		 */
		style = new Style();
		style.setStyleId("ReferenceHead");
		style.setName(new Name());
		style.getName().setVal("ReferenceHead");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setSmallCaps(new BooleanDefaultTrue());	// smallCaps
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCEHEAD, style);
		
		/*
		 * Figure
		 * Times New Roman 
		 * 8-point 
		 * spacing before="80" w:after="200"
		 */
		style = new Style();
		style.setStyleId("figurecaption");
		style.setName(new Name());
		style.getName().setVal("figurecaption");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("80"));
		style.getPPr().getSpacing().setAfter(new BigInteger("200"));
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
//		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FIGURE, style);
		
		/*
		 * tablehead
		 * spacing w:before="240" w:after="120" w:line="216" w:lineRule="auto"
		 * jc w:val="center"
		 * smallCaps
		 * sz w:val="16"
		 * rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
		 */
		style = new Style();
		style.setStyleId("tablehead");
		style.setName(new Name());
		style.getName().setVal("tablehead");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("240"));
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		style.getPPr().getSpacing().setLine(new BigInteger("216"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.AUTO);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setSmallCaps(new BooleanDefaultTrue());	// smallCaps
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TABLEHEADER, style);
		
		/*
		 * tablecolhead
		 * bold
		 * 8-point
		 * basedOn Standard
		 */
		style = new Style();
		style.setStyleId("tablecolhead");
		style.setName(new Name());
		style.getName().setVal("tablecolhead");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TABLECOLHEADER, style);
		
		/*
		 * head1
		 * Times New Roman 
		 * 12-point 
		 * small-capitals 
		 * center
		 */
		style = new Style();
		style.setStyleId("Heading1");
		style.setName(new Name());
		style.getName().setVal("Heading1");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("160"));
		style.getPPr().getSpacing().setAfter(new BigInteger("80"));
		style.getPPr().setKeepLines(new BooleanDefaultTrue());		// keeplines
		style.getPPr().setKeepNext(new BooleanDefaultTrue());		// keepnext
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setSmallCaps(new BooleanDefaultTrue());	// smallCaps
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("0"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING1, style);
		
		/*
		 * head2
		 * Times New Roman 
		 * 10-point 
		 * left
		 * italic
		 */
		style = new Style();
		style.setStyleId("Heading2");
		style.setName(new Name());
		style.getName().setVal("Heading2");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("120"));
		style.getPPr().getSpacing().setAfter(new BigInteger("60"));
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("1"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING2, style);
		
		/*
		 * head3
		 * Times New Roman 
		 * 10-point 
		 * both
		 * italic
		 */
		style = new Style();
		style.setStyleId("Heading3");
		style.setName(new Name());
		style.getName().setVal("Heading3");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("2"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING3, style);
		
		/*
		 * head4
		 * Times New Roman 
		 * 10-point 
		 * both
		 * italic
		 */
		style = new Style();
		style.setStyleId("Heading4");
		style.setName(new Name());
		style.getName().setVal("Heading4");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("40"));
		style.getPPr().getSpacing().setAfter(new BigInteger("40"));
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("3"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING4, style);
		
		/*
		 * head5
		 * Times New Roman 
		 * 10-point 
		 * both
		 * italic
		 */
		style = new Style();
		style.setStyleId("Heading5");
		style.setName(new Name());
		style.getName().setVal("Heading5");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("160"));
		style.getPPr().getSpacing().setAfter(new BigInteger("80"));
		style.getPPr().setTabs(new Tabs());
		CTTabStop tab = new CTTabStop();
		tab.setVal(STTabJc.LEFT);
		tab.setPos(new BigInteger("360"));
		style.getPPr().getTabs().getTab().add(tab);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING5, style);
		
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
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
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
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE)!=null
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getStyleMap()!=null){
				styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getStyleMap();
			} else {
				
			}
		} else {
			AcademicStylePool stylePool = new AcademicStylePool();
			stylePool.loadFile(file);
			styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.IEEE).getStyleMap();
		}
	}

}
