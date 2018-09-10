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
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Style.BasedOn;
import org.docx4j.wml.Style.Name;

import base.AppEnvironment;
import base.StaticConstants;
import db.AcademicFormatStructureDefinition;
import db.AcademicStylePool;
import lib.AbstractPublishingStyle;
import lib.ooxml.OOXMLConstantString;

public class OOXMLACMStyle implements AbstractPublishingStyle{

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

	public OOXMLACMStyle(){
		// TODO: 	1. Between Headers without spacing
		//			2. above text with spacing
		//			comments: ACM has only 4 lvl subsections
		styleList = new ArrayList<Style>();
		
		styleMap = new HashMap<String, Style>();
		
		Style style = new Style();
		
		/*
		 * 3.1	Normal or Body Text
		 * 9-points
		 * Times Roman
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
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_DEFAULTPARAGRAPH, style);
		
		/*
		 * Title
		 * Helvetica 
		 * bold
		 * 18 points
		 */
		style = new Style();
		style.setStyleId("docTitle");
		style.setName(new Name());
		style.getName().setVal("docTitle");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("120"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Helvetica");	// Helvetica
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("36"));// 18-points
		style.getRPr().setB(new BooleanDefaultTrue());		// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TITLE, style);
		
		/*
		 * Authors		
		 * Helvetica 
		 * 12 pointss
		 */
		style = new Style();
		style.setStyleId("Authors");
		style.setName(new Name());
		style.getName().setVal("Authors");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("0"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Helvetica");	// Helvetica
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));// 12-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AUTHORS, style);
		
		/*
		 * Affiliations	
		 * Helvetica
		 * 10 points	
		 */
		style = new Style();
		style.setStyleId("Affiliations");
		style.setName(new Name());
		style.getName().setVal("Affiliations");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("0"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Helvetica");	// Helvetica
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));// 10-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AFFILIATIONS, style);
		
		/*
		 * phone number
		 * Helvetica
		 * 10 points
		 */
		style = new Style();
		style.setStyleId("Phone");
		style.setName(new Name());
		style.getName().setVal("Phone");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("0"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Helvetica");	// Helvetica
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));// 10-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_PHONE, style);
		
		/*
		 * e-mail
		 * Helvetica 
		 * 12 pointss	
		 */
		style = new Style();
		style.setStyleId("Email");
		style.setName(new Name());
		style.getName().setVal("Email");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("0"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Helvetica");	// Helvetica
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));// 12-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_EMAIL, style);
		
		/*
		 * Footnotes 
		 * Times New Roman 
		 * 9-point
		 * fullwidth of page
		 */
		style = new Style();
		style.setStyleId("Footnotes");
		style.setName(new Name());
		style.getName().setVal("Footnotes");
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
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FOOTNOTES, style);
		
		/*
		 * References
		 * Times New Roman 
		 * 9-point
		 * fullwidth of column
		 */
		style = new Style();
		style.setStyleId("Reference");
		style.setName(new Name());
		style.getName().setVal("Reference");
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
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		style.getRPr().setU(null);
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCE, style);
		
		/*
		 * ReferenceHead
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
		style.getRPr().getSz().setVal(new BigInteger("24"));	// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCEHEAD, style);
		
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
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FIGURE, style);
		
		/*
		 * Caption
		 * Times New Roman 
		 * 9-point 
		 * bold	
		 */
		style = new Style();
		style.setStyleId("Caption");
		style.setName(new Name());
		style.getName().setVal("Caption");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

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
		styleMap.put(AbstractPublishingStyle.STYLENAME_CAPTION, style);
		
		/*
		 * head1
		 * Times New Roman 
		 * 12-point 
		 * bold 
		 * all-capitals 
		 * flush left		
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
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));	// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setCaps(new BooleanDefaultTrue());		// all-capitals
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("0"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING1, style);
		
		/*
		 * heading 2
		 * Times New Roman 
		 * 12-point 
		 * bold 
		 * only the initial letters capitalized (default)
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
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));	// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("1"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING2, style);
		
		/*
		 * heading 3
		 * Times New Roman 
		 * 11-point 
		 * italic 
		 * only the initial letters capitalized (default)
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
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("22"));	// 11-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("2"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING3, style);
		
		/*
		 * heading 3 char
		 */
		style = new Style();
		style.setStyleId("Heading3Char");
		style.setName(new Name());
		style.getName().setVal("Heading3Char");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("22"));	// 11-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING3CHAR, style);
		
		/*
		 * heading 4
		 * Times New Roman 
		 * 11-point 
		 * italic 
		 * only the initial letters capitalized (default)
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
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("22"));	// 11-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("3"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING4, style);
		/*
		 * head1
		 * Times New Roman 
		 * 12-point 
		 * bold 
		 * all-capitals 
		 * flush left		
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
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));	// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setCaps(new BooleanDefaultTrue());		// all-capitals
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("0"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING5, style);
		/*
		 * abstract head
		 * Times New Roman 
		 * 12-point 
		 * bold 
		 * all-capitals 
		 * flush left		
		 */
		style = new Style();
		style.setStyleId("AbstractHeader");
		style.setName(new Name());
		style.getName().setVal("AbstractHeader");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));	// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setCaps(new BooleanDefaultTrue());		// all-capitals
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_ABSTRACTHEADER, style);
		
		/*
		 * abstract
		 * Times New Roman 
		 * 9-point 
		 * normal 
		 * both		
		 */
		style = new Style();
		style.setStyleId("Abstract");
		style.setName(new Name());
		style.getName().setVal("Abstract");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);		// flush both
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_ABSTRACT, style);
		
		/*
		 * Keyword head
		 * Times New Roman 
		 * 12-point 
		 * bold 
		 * all-capitals 
		 * flush left		
		 */
		style = new Style();
		style.setStyleId("KeywordHeader");
		style.setName(new Name());
		style.getName().setVal("KeywordHeader");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("24"));	// 12-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setCaps(new BooleanDefaultTrue());		// all-capitals
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_KEYWORDHEAD, style);
		
		/*
		 * keyword
		 * Times New Roman 
		 * 9-point 
		 * normal 
		 * both		
		 */
		style = new Style();
		style.setStyleId("Keyword");
		style.setName(new Name());
		style.getName().setVal("Keyword");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);		// flush both
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_KEYWORD, style);
		
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
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM)!=null
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap()!=null){
				styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap();
			} else {
				
			}
		} else {
			AcademicStylePool stylePool = new AcademicStylePool();
			stylePool.loadFile(file);
			styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap();
		}
	}
}
