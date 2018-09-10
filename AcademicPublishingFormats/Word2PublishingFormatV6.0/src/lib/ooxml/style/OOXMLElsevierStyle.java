package lib.ooxml.style;

import java.io.File;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.PPr;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STTabJc;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tabs;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.Style.BasedOn;
import org.docx4j.wml.Style.Name;

import base.AppEnvironment;
import base.StaticConstants;
import db.AcademicFormatStructureDefinition;
import db.AcademicStylePool;
import lib.AbstractPublishingStyle;
import lib.ooxml.OOXMLConstantString;

public class OOXMLElsevierStyle implements AbstractPublishingStyle {

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
//		this.styleList = new ArrayList<Style>();
//		for(Map.Entry<String, Style> entry : styleMap.entrySet()){
//			this.styleList.add(entry.getValue());
//		}
	}

	public OOXMLElsevierStyle() {
		// TODO Auto-generated constructor stub
		styleList = new ArrayList<Style>();
		
		styleMap = new HashMap<String, Style>();
		
		Style style = new Style();
		
		/*
		 * Normal or Body Text
		 * 10-points
		 * Times New Roman
		 */
		style = new Style();
		
		style.setStyleId("Els-body-text");
		style.setName(new Name());
		style.getName().setVal("Els-body-text");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
		style.getPPr().getInd().setFirstLine(new BigInteger("238"));
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_DEFAULTPARAGRAPH, style);
		
		/*
		 * Normal or Body Text
		 * 10-points
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
		style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
		style.getPPr().getInd().setFirstLine(new BigInteger("238"));
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_DEFAULTPARAGRAPH, style);
		
		/*
		 * Title
		 * Times New Roman 
		 * center
		 * 17 points
		 */
		style = new Style();
		style.setStyleId("Els-Title");
		style.setName(new Name());
		style.getName().setVal("Els-Title");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setAutoRedefine(new BooleanDefaultTrue());
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-Author");

		style.setPPr(new PPr());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("240"));
		style.getPPr().getSpacing().setLine(new BigInteger("400"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("34"));// 17-points
//		style.getRPr().setB(new BooleanDefaultTrue());		// bold
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_TITLE, style);
		
		/*
		 * Authors		
		 * Times New Roman 
		 * 13 points
		 */
		style = new Style();
		style.setStyleId("Els-Author");
		style.setName(new Name());
		style.getName().setVal("Els-Author");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("160"));
		style.getPPr().getSpacing().setLine(new BigInteger("300"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		
		style.setRPr(new RPr());
		style.getRPr().setNoProof(new BooleanDefaultTrue());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("26"));// 13-points
		style.getRPr().setLang(new CTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AUTHORS, style);
		
		/*
		 * Affiliations	
		 * Times New Roman
		 * 8 points	
		 */
		style = new Style();
		style.setStyleId("Els-Affiliation");
		style.setName(new Name());
		style.getName().setVal("Els-Affiliation");
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-Abstract-head");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.CENTER);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));// 8-points
		style.getRPr().setI(new BooleanDefaultTrue());
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_AFFILIATIONS, style);
		
		/*
		 * Footnotes 
		 * Times New Roman 
		 * 8-point
		 * 
		 */
		style = new Style();
		style.setStyleId("Els-footnote");
		style.setName(new Name());
		style.getName().setVal("Els-footnote");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
		style.getPPr().getInd().setFirstLine(new BigInteger("240"));
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FOOTNOTES, style);
		
		/*
		 * References
		 * Times New Roman 
		 * 9-point
		 * fullwidth of column
		 */
		style = new Style();
		style.setStyleId("Els-reference");
		style.setName(new Name());
		style.getName().setVal("Els-reference");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(new PPr());
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
		style.getPPr().getInd().setLeft(new BigInteger("312"));
		style.getPPr().getInd().setHanging(new BigInteger("312"));
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
		style.getRPr().setNoProof(new BooleanDefaultTrue());
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCE, style);
		
		/*
		 * ReferenceHead
		 */
		style = new Style();
		style.setStyleId("Els-reference-head");
		style.setName(new Name());
		style.getName().setVal("Els-reference-head");
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-reference");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		
		style.setPPr(Context.getWmlObjectFactory().createPPr());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("200"));
		style.getPPr().getSpacing().setBefore(new BigInteger("480"));
		style.getPPr().getSpacing().setLine(new BigInteger("220"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_REFERENCEHEAD, style);
		
		/*
		 * FigureCaption	// not defined in template
		 * Times New Roman 
		 * 8-point 
		 */
		style = new Style();
		style.setStyleId("Figure");
		style.setName(new Name());
		style.getName().setVal("Figure");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("240"));
		style.getPPr().getSpacing().setBefore(new BigInteger("200"));
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_FIGURE, style);
		
		/*
		 * Caption
		 * Times New Roman 
		 * 8-point 
		 */
		style = new Style();
		style.setStyleId("Els-caption");
		style.setName(new Name());
		style.getName().setVal("Els-caption");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setKeepLines(new BooleanDefaultTrue());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("240"));
		style.getPPr().getSpacing().setBefore(new BigInteger("200"));
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_CAPTION, style);
		
		/*
		 * head1
		 * Times New Roman 
		 * 10-point 
		 * bold 
		 * flush left		
		 */
		style = new Style();
		style.setStyleId("Els-1storder-head");
		style.setCustomStyle(true);
		style.setName(new Name());
		style.getName().setVal("Els-1storder-head");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-body-text");

		style.setPPr(new PPr());
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("240"));
		style.getPPr().getSpacing().setBefore(new BigInteger("240"));
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
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
		 * 10-point 
		 * italic 
		 * only the initial letters capitalized (default)
		 */
		style = new Style();
		style.setStyleId("Els-2ndorder-head");
		style.setName(new Name());
		style.getName().setVal("Els-2ndorder-head");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setCustomStyle(true);
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-body-text");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("240"));
		style.getPPr().getSpacing().setBefore(new BigInteger("240"));
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
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
		 * 10-point 
		 * italic 
		 */
		style = new Style();
		style.setStyleId("Els-3rdorder-head");
		style.setName(new Name());
		style.getName().setVal("Els-3rdorder-head");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setCustomStyle(true);
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-body-text");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("240"));
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
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
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING3CHAR, style);
		
		/*
		 * heading 4
		 * Times New Roman 
		 * 10-point 
		 * italic 
		 */
		style = new Style();
		style.setStyleId("Els-4thorder-head");
		style.setName(new Name());
		style.getName().setVal("Els-4thorder-head");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);
		style.setCustomStyle(true);
		style.setNext(Context.getWmlObjectFactory().createStyleNext());
		style.getNext().setVal("Els-body-text");
		
		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("240"));
		style.getPPr().getSpacing().setLine(new BigInteger("240"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("22"));	// 11-points
		style.getRPr().setI(new BooleanDefaultTrue());			// italic
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		style.getPPr().setNumPr(new NumPr());
		style.getPPr().getNumPr().setNumId(new NumId());
		style.getPPr().getNumPr().getNumId().setVal(new BigInteger(StaticConstants.OOXML_SECTION_HEADER_NUMBERING_ABSTRACTNUMID));
		style.getPPr().getNumPr().setIlvl(new Ilvl());
		style.getPPr().getNumPr().getIlvl().setVal(new BigInteger("3"));
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING4, style);
		
		/*
		 * head5			// for Els-acknowledgement && Els-appendixhead
		 * Times New Roman 
		 * 10-point 
		 * both
		 * bold
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
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);
		style.getPPr().setSpacing(new Spacing());
		style.getPPr().getSpacing().setBefore(new BigInteger("480"));
		style.getPPr().getSpacing().setAfter(new BigInteger("240"));
		style.getPPr().getSpacing().setLine(new BigInteger("220"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setTabs(new Tabs());
		CTTabStop tab = new CTTabStop();
		tab.setVal(STTabJc.LEFT);
//		tab.setPos(new BigInteger("360"));
//		style.getPPr().getTabs().getTab().add(tab);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().getRFonts().setHAnsi("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("20"));	// 10-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_HEADING5, style);
		
		/*
		 * abstract head
		 * Times New Roman 
		 * 9-point 
		 * bold 
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
//		style.getPPr().setJc(new Jc());
//		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		style.getPPr().setKeepNext(new BooleanDefaultTrue());
		style.getPPr().setSuppressAutoHyphens(new BooleanDefaultTrue());
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("220"));
		style.getPPr().getSpacing().setLine(new BigInteger("220"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
		style.getPPr().getPBdr().setTop(Context.getWmlObjectFactory().createCTBorder());
		style.getPPr().getPBdr().getTop().setVal(STBorder.SINGLE);
		style.getPPr().getPBdr().getTop().setSz(new BigInteger("4"));
		style.getPPr().getPBdr().getTop().setSpace(new BigInteger("10"));
		style.getPPr().getPBdr().getTop().setColor("auto");
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		style.getRPr().setB(new BooleanDefaultTrue());			// bold
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
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
		style.setStyleId("Els-Abstract-text");
		style.setName(new Name());
		style.getName().setVal("Els-Abstract-text");
		style.setType(OOXMLConstantString.STYLE_TYPE_PARAGRAPH);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);		// flush both
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setLine(new BigInteger("220"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 9-points
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
		styleList.add(style);
		styleMap.put(AbstractPublishingStyle.STYLENAME_ABSTRACT, style);
		
		/*
		 * Keyword head
		 * Times New Roman 
		 * 8-point 
		 * italic 
		 * flush left		
		 */
		style = new Style();
		style.setStyleId("KeywordHeader");
		style.setName(new Name());
		style.getName().setVal("KeywordHeader");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.LEFT);		// flush left
		style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
		style.getPPr().getPBdr().setBottom(Context.getWmlObjectFactory().createCTBorder());
		style.getPPr().getPBdr().getBottom().setVal(STBorder.SINGLE);
		style.getPPr().getPBdr().getBottom().setSz(new BigInteger("4"));
		style.getPPr().getPBdr().getBottom().setSpace(new BigInteger("10"));
		style.getPPr().getPBdr().getBottom().setColor("auto");
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("200"));
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setI(new BooleanDefaultTrue());			// bold
		style.getRPr().setNoProof(new BooleanDefaultTrue());
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
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
		style.setStyleId("Els-keywords");
		style.setName(new Name());
		style.getName().setVal("Els-keywords");
		style.setType(OOXMLConstantString.STYLE_TYPE_CHARACTER);
		style.setBasedOn(new BasedOn());
		style.getBasedOn().setVal(OOXMLConstantString.STYLE_DEFAULT_PARAGRAPH_FONT);

		style.setPPr(new PPr());
		style.getPPr().setJc(new Jc());
		style.getPPr().getJc().setVal(JcEnumeration.BOTH);		// flush both
		style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
		style.getPPr().getPBdr().setBottom(Context.getWmlObjectFactory().createCTBorder());
		style.getPPr().getPBdr().getBottom().setVal(STBorder.SINGLE);
		style.getPPr().getPBdr().getBottom().setSz(new BigInteger("4"));
		style.getPPr().getPBdr().getBottom().setSpace(new BigInteger("10"));
		style.getPPr().getPBdr().getBottom().setColor("auto");
		style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		style.getPPr().getSpacing().setAfter(new BigInteger("200"));
		style.getPPr().getSpacing().setLine(new BigInteger("200"));
		style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
		
		style.setRPr(new RPr());
		style.getRPr().setColor(new Color());
		style.getRPr().getColor().setVal("000000");
		style.getRPr().setRFonts(new RFonts());
		style.getRPr().getRFonts().setAscii("Times New Roman");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("16"));	// 8-points
		style.getRPr().setNoProof(new BooleanDefaultTrue());
		style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
		style.getRPr().getLang().setVal("en-US");
		
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
		style.getRPr().getRFonts().setAscii("Arial");	// Times New Roman
		style.getRPr().setSz(new HpsMeasure());
		style.getRPr().getSz().setVal(new BigInteger("18"));	// 8-points
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
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ELSEVIER)!=null
					&& AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ELSEVIER).getStyleMap()!=null){
				styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ELSEVIER).getStyleMap();
			} else {
				
			}
		} else {
			AcademicStylePool stylePool = new AcademicStylePool();
			stylePool.loadFile(file);
			styleMap = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ELSEVIER).getStyleMap();
		}
	}

}
