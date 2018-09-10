package db;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTColumns;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.Color;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.PPrBase.PBdr;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STShd;
import org.docx4j.wml.STTabJc;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.Style;
import org.docx4j.wml.Style.BasedOn;
import org.docx4j.wml.Style.Name;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import base.AppEnvironment;
import lib.ooxml.OOXMLConstantString;
import tools.LogUtils;

public class AcademicStylePool {

	private Map<String, AcademicStyleMap> styles;
	
	public Map<String, AcademicStyleMap> getStyles() {
		return styles;
	}

	public void setStyles(Map<String, AcademicStyleMap> styles) {
		this.styles = styles;
	}

	public AcademicStylePool() {
		// TODO Auto-generated constructor stub
		this.styles = new HashMap<String, AcademicStyleMap>();
		
	}
	
	public void loadDetails(AcademicStyleMap styleMap, JSONObject jsonObjItem, Style style, String keyStr){
		if(jsonObjItem!=null){
			if(jsonObjItem.get(AcademicFormatStructureDefinition.DEFAULTSYLE)!=null) {
				style.setDefault("true".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.DEFAULTSYLE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.ID)!=null){
				style.setStyleId((String) jsonObjItem.get(AcademicFormatStructureDefinition.ID));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.NAME)!=null){
				style.setName(Context.getWmlObjectFactory().createStyleName());
				style.getName().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.NAME));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.TYPE)!=null){
				style.setType((String)jsonObjItem.get(AcademicFormatStructureDefinition.TYPE));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARAKEEPLINE)!=null){
				if("true".equals(((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAKEEPLINE)).toLowerCase())){
					style.getPPr().setKeepLines(new BooleanDefaultTrue());
				} else {
					style.getPPr().setKeepLines(null);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARAKEEPNEXT)!=null){
				if("true".equals(((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAKEEPNEXT)).toLowerCase())){
					style.getPPr().setKeepNext(new BooleanDefaultTrue());
				} else {
					style.getPPr().setKeepNext(null);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARANEXT)!=null){
				style.setNext(Context.getWmlObjectFactory().createStyleNext());
				style.getNext().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARANEXT));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTSIZE)!=null){
//				style.getPPr().getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
//				style.getPPr().getRPr().getSz().setVal(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.FONTSIZE)));
				style.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
				style.getRPr().getSz().setVal(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.FONTSIZE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTFAMILY)!=null){
//				style.getPPr().getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
//				style.getPPr().getRPr().getLang().setVal((String) jsonObjItem.get(AcademicFormatStructureDefinition.FONTFAMILY));
//				style.getRPr().setLang(Context.getWmlObjectFactory().createCTLanguage());
//				style.getRPr().getLang().setVal((String) jsonObjItem.get(AcademicFormatStructureDefinition.FONTFAMILY));
				style.getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
				style.getRPr().getRFonts().setAscii((String) jsonObjItem.get(AcademicFormatStructureDefinition.FONTFAMILY));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTCOLOR)!=null){
//				style.getPPr().getRPr().setColor(Context.getWmlObjectFactory().createColor());
//				style.getPPr().getRPr().getColor().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTCOLOR));
				style.getRPr().setColor(Context.getWmlObjectFactory().createColor());
				style.getRPr().getColor().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTCOLOR));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARAJC)!=null){
				style.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAJC)).toLowerCase().equals("center")){
					style.getPPr().getJc().setVal(JcEnumeration.CENTER);
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAJC)).toLowerCase().equals("left")){
					style.getPPr().getJc().setVal(JcEnumeration.LEFT);
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAJC)).toLowerCase().equals("right")){
					style.getPPr().getJc().setVal(JcEnumeration.RIGHT);
				} else {
					style.getPPr().getJc().setVal(JcEnumeration.BOTH);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.BASEDON)!=null){
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.BASEDON)).toLowerCase().equals("default")){
					style.setBasedOn(Context.getWmlObjectFactory().createStyleBasedOn());
					style.getBasedOn().setVal(AcademicFormatStructureDefinition.NORMAL);
				} else {
					style.setBasedOn(Context.getWmlObjectFactory().createStyleBasedOn());
					style.getBasedOn().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.BASEDON));
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.LINK)!=null){
				style.setLink(Context.getWmlObjectFactory().createStyleLink());
				style.getLink().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.LINK));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTBOLD)!=null){
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTBOLD)).toLowerCase().equals("true")){
//					style.getPPr().getRPr().setB(new BooleanDefaultTrue());
					style.getRPr().setB(new BooleanDefaultTrue());
					style.getRPr().getB().setVal(true);
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTBOLD)).toLowerCase().equals("false")) {
					style.getRPr().setB(new BooleanDefaultTrue());
					style.getRPr().getB().setVal(false);
				} else {
//					style.getPPr().getRPr().setB(null);
					style.getRPr().setB(null);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTITATLIC)!=null){
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTITATLIC)).toLowerCase().equals("true")){
//					style.getPPr().getRPr().setI(new BooleanDefaultTrue());
					style.getRPr().setI(new BooleanDefaultTrue());
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTITATLIC)).toLowerCase().equals("false")) {
					style.getRPr().setI(new BooleanDefaultTrue());
					style.getRPr().getI().setVal(false);
				} else {
//					style.getPPr().getRPr().setI(null);
					style.getRPr().setI(null);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTCAPS)!=null){
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTCAPS)).toLowerCase().equals("true")){
//					style.getPPr().getRPr().setCaps(new BooleanDefaultTrue());
					style.getRPr().setCaps(new BooleanDefaultTrue());
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTCAPS)).toLowerCase().equals("false")) {
					style.getRPr().setCaps(new BooleanDefaultTrue());
					style.getRPr().getCaps().setVal(false);
				} else {
//					style.getPPr().getRPr().setCaps(null);
					style.getRPr().setCaps(null);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.FONTSMALLCAPS)!=null){
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTSMALLCAPS)).toLowerCase().equals("true")){
//					style.getPPr().getRPr().setSmallCaps(new BooleanDefaultTrue());
					style.getRPr().setSmallCaps(new BooleanDefaultTrue());
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.FONTSMALLCAPS)).toLowerCase().equals("false")) {
					style.getRPr().setSmallCaps(new BooleanDefaultTrue());
					style.getRPr().getSmallCaps().setVal(false);
				} else {
//					style.getPPr().getRPr().setSmallCaps(null);
					style.getRPr().setSmallCaps(null);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARAAFTER)!=null){
				if(style.getPPr().getSpacing()==null){
					style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
				}
				style.getPPr().getSpacing().setAfter(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAAFTER)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABEFORE)!=null){
				if(style.getPPr().getSpacing()==null){
					style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
				}
				style.getPPr().getSpacing().setAfter(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABEFORE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARALINE)!=null){
				if(style.getPPr().getSpacing()==null){
					style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
				}
				style.getPPr().getSpacing().setLine(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARALINE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARALINERULE)!=null){
				if(style.getPPr().getSpacing()==null){
					style.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
				}
				if("auto".equals(((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARALINERULE)).toLowerCase())){
					style.getPPr().getSpacing().setLineRule(STLineSpacingRule.AUTO);
				} else if("exact".equals(((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARALINERULE)).toLowerCase())){
					style.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
				} else if("atleast".equals(((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARALINERULE)).toLowerCase())){
					style.getPPr().getSpacing().setLineRule(STLineSpacingRule.AT_LEAST);
				} else {
					style.getPPr().getSpacing().setLineRule(STLineSpacingRule.AUTO);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARALEFT)!=null){
				if(style.getPPr().getInd()==null){
					style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
				}
				style.getPPr().getInd().setLeft(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARALEFT)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARARIGHT)!=null){
				if(style.getPPr().getInd()==null){
					style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
				}
				style.getPPr().getInd().setRight(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARARIGHT)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARAFIRSTLINE)!=null){
				if(style.getPPr().getInd()==null){
					style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
				}
				style.getPPr().getInd().setFirstLine(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAFIRSTLINE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARAHANGING)!=null){
				if(style.getPPr().getInd()==null){
					style.getPPr().setInd(Context.getWmlObjectFactory().createPPrBaseInd());
				}
				style.getPPr().getInd().setHanging(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARAHANGING)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGSZW)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgSz()==null){
					style.getPPr().getSectPr().setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
				}
				style.getPPr().getSectPr().getPgSz().setW(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGSZW)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGSZH)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgSz()==null){
					style.getPPr().getSectPr().setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
				}
				style.getPPr().getSectPr().getPgSz().setH(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGSZH)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGSZCODE)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgSz()==null){
					style.getPPr().getSectPr().setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
				}
				style.getPPr().getSectPr().getPgSz().setCode(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGSZCODE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINBOTTOM)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setBottom(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINBOTTOM)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINTOP)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setTop(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINTOP)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINLEFT)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setLeft(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINLEFT)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINRIGHT)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setRight(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGMARGINRIGHT)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGGUTTER)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setGutter(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGGUTTER)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGHEADER)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setHeader(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGHEADER)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGFOOTER)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getPgMar()==null){
					style.getPPr().getSectPr().setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				}
				style.getPPr().getSectPr().getPgMar().setFooter(new BigInteger((String) jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTPGFOOTER)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTCOLNUM)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getCols()==null){
					style.getPPr().getSectPr().setCols(Context.getWmlObjectFactory().createCTColumns());
				}
				style.getPPr().getSectPr().getCols().setNum(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTCOLNUM)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTCOLSPACING)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getCols()==null){
					style.getPPr().getSectPr().setCols(Context.getWmlObjectFactory().createCTColumns());
				}
				style.getPPr().getSectPr().getCols().setSpace(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTCOLSPACING)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTLINEPITCH)!=null){
				if(style.getPPr().getSectPr()==null){
					style.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
				}
				if(style.getPPr().getSectPr().getDocGrid()==null){
					style.getPPr().getSectPr().setDocGrid(Context.getWmlObjectFactory().createCTDocGrid());
				}
				style.getPPr().getSectPr().getDocGrid().setLinePitch(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARASECTLINEPITCH)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARANUMLIST)!=null){
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARANUMLIST)).toLowerCase().equals("null")){
					style.getPPr().setNumPr(null);
				} else {
					if(style.getPPr().getNumPr()==null){
						style.getPPr().setNumPr(Context.getWmlObjectFactory().createPPrBaseNumPr());
					}
					style.getPPr().getNumPr().setNumId(Context.getWmlObjectFactory().createPPrBaseNumPrNumId());
					style.getPPr().getNumPr().getNumId().setVal(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARANUMLIST)));
					AbstractNum abstNum = getAbstractNum(keyStr, styleMap);
					abstNum.setAbstractNumId(style.getPPr().getNumPr().getNumId().getVal());
				}
			} else {
				style.getPPr().setNumPr(null);
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARALVL)!=null){
				if(style.getPPr().getNumPr()==null){
					style.getPPr().setNumPr(Context.getWmlObjectFactory().createPPrBaseNumPr());
				}
				style.getPPr().getNumPr().setIlvl(Context.getWmlObjectFactory().createPPrBaseNumPrIlvl());
				style.getPPr().getNumPr().getIlvl().setVal(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARALVL)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPTYPE)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getTop()==null){
					style.getPPr().getPBdr().setTop(Context.getWmlObjectFactory().createCTBorder());
				}
				if("single".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPTYPE))){
					style.getPPr().getPBdr().getTop().setVal(STBorder.SINGLE);
				} else if("double".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPTYPE))){
					style.getPPr().getPBdr().getTop().setVal(STBorder.DOUBLE);
				} else if("dot_dash".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPTYPE))){
					style.getPPr().getPBdr().getTop().setVal(STBorder.DOT_DASH);
				} else {
					style.getPPr().getPBdr().getTop().setVal(STBorder.SINGLE);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPSZ)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getTop()==null){
					style.getPPr().getPBdr().setTop(Context.getWmlObjectFactory().createCTBorder());
				}
				style.getPPr().getPBdr().getTop().setSz(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPSZ)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPSPACE)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getTop()==null){
					style.getPPr().getPBdr().setTop(Context.getWmlObjectFactory().createCTBorder());
				}
				style.getPPr().getPBdr().getTop().setSpace(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPSPACE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPCOLOR)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getTop()==null){
					style.getPPr().getPBdr().setTop(Context.getWmlObjectFactory().createCTBorder());
				}
				style.getPPr().getPBdr().getTop().setColor((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRTOPCOLOR));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMTYPE)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getBottom()==null){
					style.getPPr().getPBdr().setBottom(Context.getWmlObjectFactory().createCTBorder());
				}
				if("single".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMTYPE))){
					style.getPPr().getPBdr().getBottom().setVal(STBorder.SINGLE);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMSZ)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getBottom()==null){
					style.getPPr().getPBdr().setBottom(Context.getWmlObjectFactory().createCTBorder());
				}
				style.getPPr().getPBdr().getBottom().setSz(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMSZ)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMSPACE)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getBottom()==null){
					style.getPPr().getPBdr().setBottom(Context.getWmlObjectFactory().createCTBorder());
				}
				style.getPPr().getPBdr().getBottom().setSpace(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMSPACE)));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMCOLOR)!=null){
				if(style.getPPr().getPBdr()==null){
					style.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
				}
				if(style.getPPr().getPBdr().getBottom()==null){
					style.getPPr().getPBdr().setBottom(Context.getWmlObjectFactory().createCTBorder());
				}
				style.getPPr().getPBdr().getBottom().setColor((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARABDRBOTTOMCOLOR));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.SHD_VAL)!=null){
				if(style.getPPr().getShd()==null){
					style.getPPr().setShd(Context.getWmlObjectFactory().createCTShd());
				}
				if("clear".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.SHD_VAL))) {
					style.getPPr().getShd().setVal(STShd.CLEAR);
				} else if("nil".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.SHD_VAL))){
					style.getPPr().getShd().setVal(STShd.NIL);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.SHD_COLOR)!=null){
				if(style.getPPr().getShd()==null){
					style.getPPr().setShd(Context.getWmlObjectFactory().createCTShd());
				}
				style.getPPr().getShd().setColor((String)jsonObjItem.get(AcademicFormatStructureDefinition.SHD_COLOR));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.SHD_FILL)!=null){
				if(style.getPPr().getShd()==null){
					style.getPPr().setShd(Context.getWmlObjectFactory().createCTShd());
				}
				style.getPPr().getShd().setFill((String)jsonObjItem.get(AcademicFormatStructureDefinition.SHD_FILL));
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARATABVAL)!=null){
				if(style.getPPr().getTabs()==null){
					style.getPPr().setTabs(Context.getWmlObjectFactory().createTabs());
					CTTabStop tab = Context.getWmlObjectFactory().createCTTabStop();
					style.getPPr().getTabs().getTab().add(tab);
				}
				if(style.getPPr().getTabs().getTab()!=null && style.getPPr().getTabs().getTab().size()>0){
					boolean isDone = false;
					for(CTTabStop tab : style.getPPr().getTabs().getTab()){
						if(tab.getVal()!=null){
							
						} else {
							if("left".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARATABVAL))){
								tab.setVal(STTabJc.LEFT);
								isDone = true;
							} else if("center".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARATABVAL))){
								tab.setVal(STTabJc.CENTER);
								isDone = true;
							}
						}
					}
					if(!isDone){
						CTTabStop tab = Context.getWmlObjectFactory().createCTTabStop();
						style.getPPr().getTabs().getTab().add(tab);
						if("left".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARATABVAL))){
							tab.setVal(STTabJc.LEFT);
							isDone = true;
						} else if("center".equals((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARATABVAL))){
							tab.setVal(STTabJc.CENTER);
							isDone = true;
						}
					}
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.PARATABPOS)!=null){
				if(style.getPPr().getTabs()==null){
					style.getPPr().setTabs(Context.getWmlObjectFactory().createTabs());
					CTTabStop tab = Context.getWmlObjectFactory().createCTTabStop();
					style.getPPr().getTabs().getTab().add(tab);
				}
				if(style.getPPr().getTabs().getTab()!=null && style.getPPr().getTabs().getTab().size()>0){
					boolean isDone = false;
					for(CTTabStop tab : style.getPPr().getTabs().getTab()){
						if(tab.getPos()!=null){
							
						} else {
							tab.setPos(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARATABPOS)));
						}
					}
					if(!isDone){
						CTTabStop tab = Context.getWmlObjectFactory().createCTTabStop();
						style.getPPr().getTabs().getTab().add(tab);
						tab.setPos(new BigInteger((String)jsonObjItem.get(AcademicFormatStructureDefinition.PARATABPOS)));
					}
				}
			}
			
			/*
			 * numbering
			 */
			if(jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGFMT)!=null){
				Lvl lvl = getLvl(keyStr, styleMap);
				
				if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGFMT)).toLowerCase().equals("decimal")){
					lvl.getNumFmt().setVal(NumberFormat.DECIMAL);
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGFMT)).toLowerCase().equals("upperletter")){
					lvl.getNumFmt().setVal(NumberFormat.UPPER_LETTER);
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGFMT)).toLowerCase().equals("lowerletter")){
					lvl.getNumFmt().setVal(NumberFormat.LOWER_LETTER);
				} else if(((String)jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGFMT)).toLowerCase().equals("upperroman")){
					lvl.getNumFmt().setVal(NumberFormat.UPPER_ROMAN);
				} else {
					lvl.getNumFmt().setVal(NumberFormat.DECIMAL);
				}
			}
			if(jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGLVLTEXT)!=null){
				Lvl lvl = getLvl(keyStr, styleMap);
				lvl.setLvlText(Context.getWmlObjectFactory().createLvlLvlText());
				lvl.getLvlText().setVal((String)jsonObjItem.get(AcademicFormatStructureDefinition.NUMBERINGLVLTEXT));
			}
			/*
			 * 
			 */
			style.setCustomStyle(true);
			
			styleMap.getStyleMap().put(keyStr, style);
		}
	}
	private AbstractNum getAbstractNum(String keyStr, AcademicStyleMap styleMap){
		AbstractNum abstNum = null;
		if(keyStr.equals(AcademicFormatStructureDefinition.BIBLIOGRAPHY)){
			abstNum = styleMap.getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGREFSTRING);
		} else if(keyStr.equals(AcademicFormatStructureDefinition.FIGURE)) {
			abstNum = styleMap.getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGFIGURESTRING);
		} else if(keyStr.equals(AcademicFormatStructureDefinition.CAPTION)) {
			abstNum = styleMap.getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGCAPTIONSTRING);
		} else {
			abstNum = styleMap.getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGSECTIONSTRING);
		}
		return abstNum;
	}
	private Lvl getLvl(String keyStr, AcademicStyleMap styleMap){
		AbstractNum abstNum = getAbstractNum(keyStr, styleMap);
		
		List<Lvl> lvlList = abstNum.getLvl();
		int lvlIndex = -1;
		if(keyStr.equals(AcademicFormatStructureDefinition.HEADING1)){
			lvlIndex = 0;
		} else if(keyStr.equals(AcademicFormatStructureDefinition.HEADING2)){
			lvlIndex = 1;
		} else if(keyStr.equals(AcademicFormatStructureDefinition.HEADING3)){
			lvlIndex = 2;
		} else if(keyStr.equals(AcademicFormatStructureDefinition.HEADING4)){
			lvlIndex = 3;
		}
		Lvl lvl = null;
		for (int i = 0; i < lvlList.size(); i++) {
			if(lvlList.get(i).getIlvl()!=null && lvlList.get(i).getIlvl().intValue()==lvlIndex){
				lvl = lvlList.get(i);
				break;
			}
		}
		if(lvl == null){
			lvl = Context.getWmlObjectFactory().createLvl();
			lvl.setIlvl(BigInteger.valueOf(lvlIndex));
			lvl.setPStyle(Context.getWmlObjectFactory().createLvlPStyle());
			lvl.getPStyle().setVal(keyStr);
			lvl.setNumFmt(Context.getWmlObjectFactory().createNumFmt());
			lvl.setStart(Context.getWmlObjectFactory().createLvlStart());
			lvl.getStart().setVal(BigInteger.valueOf(1));
			lvlList.add(lvl);
		}
		return lvl;
	}
	private void autoFillAndModify(AcademicStyleMap styleMap){
		/*
		 * do some intelligent adapt 
		 */
		if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.TITLE)!=null && styleMap.getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT)!=null){
			styleMap.getStyleMap().get(AcademicFormatStructureDefinition.TITLE).getPPr().setSectPr(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT).getPPr().getSectPr());
			if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.TITLE).getPPr().getSectPr()!=null 
					&& styleMap.getStyleMap().get(AcademicFormatStructureDefinition.TITLE).getPPr().getSectPr().getCols()!=null){
				styleMap.getStyleMap().get(AcademicFormatStructureDefinition.TITLE).getPPr().getSectPr().getCols().setNum(BigInteger.valueOf(1));
			} else {
				
			}
		}
		if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.NORMALCHAR)==null){
			Style style = Context.getWmlObjectFactory().createStyle();
			style.setRPr(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.NORMAL).getRPr());
			style.setStyleId(AcademicFormatStructureDefinition.NORMALCHAR);
			style.setName(Context.getWmlObjectFactory().createStyleName());
			style.getName().setVal(AcademicFormatStructureDefinition.NORMALCHAR);
			style.setType("character");
			styleMap.getStyleMap().put(AcademicFormatStructureDefinition.NORMALCHAR, style);
		}
		if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.HEADING5)!=null && styleMap.getStyleMap().get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER)==null){
			styleMap.getStyleMap().put(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER, styleMap.getStyleMap().get(AcademicFormatStructureDefinition.HEADING5));
		}
		if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.HEADING5CHAR)==null){
			styleMap.getStyleMap().put(AcademicFormatStructureDefinition.HEADING5CHAR, styleMap.getStyleMap().get(AcademicFormatStructureDefinition.NORMALCHAR));
		}
		if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.ACKNOWLEDGMENT)==null){
			styleMap.getStyleMap().put(AcademicFormatStructureDefinition.ACKNOWLEDGMENT, styleMap.getStyleMap().get(AcademicFormatStructureDefinition.NORMAL));
		}
		if(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTCHAR)==null){
			styleMap.getStyleMap().put(AcademicFormatStructureDefinition.ACKNOWLEDGMENTCHAR, styleMap.getStyleMap().get(AcademicFormatStructureDefinition.NORMALCHAR));
		}
		
//		AbstractNum abstNum = styleMap.getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGSECTIONSTRING);
//		if(abstNum!=null){
//			if(abstNum.getAbstractNumId()!=null){
//				
//			} else {
//				abstNum.setAbstractNumId(BigInteger.valueOf(styleMap.getStyleMap().get(AcademicFormatStructureDefinition.HEADING1).getPPr().getNumPr().getNumId().getVal().intValue()));
//			}
//		}
	}
	
	public void loadFile(File styleFile){
		try {
			JSONParser parser = new JSONParser();
			Object obj = parser.parse(new FileReader(styleFile));
			JSONObject jsonRootObject = (JSONObject) obj;
			if(jsonRootObject!=null){
				AcademicStyleMap styleMap = new AcademicStyleMap();
				styleMap.setStyleMap(new HashMap<String, Style>());
				styleMap.setNumberingMap(new HashMap<String, AbstractNum>());
				AbstractNum abstNum = Context.getWmlObjectFactory().createNumberingAbstractNum();
				abstNum.setAbstractNumId(BigInteger.valueOf(99));
				abstNum.setMultiLevelType(Context.getWmlObjectFactory().createNumberingAbstractNumMultiLevelType());
				abstNum.getMultiLevelType().setVal("multilevel");
				styleMap.getNumberingMap().put(AcademicFormatStructureDefinition.NUMBERINGSECTIONSTRING, abstNum);
				
				abstNum = Context.getWmlObjectFactory().createNumberingAbstractNum();
				abstNum.setAbstractNumId(BigInteger.valueOf(98));
				abstNum.setMultiLevelType(Context.getWmlObjectFactory().createNumberingAbstractNumMultiLevelType());
				abstNum.getMultiLevelType().setVal("singleLevel");
				styleMap.getNumberingMap().put(AcademicFormatStructureDefinition.NUMBERINGREFSTRING, abstNum);
				
				abstNum = Context.getWmlObjectFactory().createNumberingAbstractNum();
				abstNum.setAbstractNumId(BigInteger.valueOf(97));
				abstNum.setMultiLevelType(Context.getWmlObjectFactory().createNumberingAbstractNumMultiLevelType());
				abstNum.getMultiLevelType().setVal("singleLevel");
				styleMap.getNumberingMap().put(AcademicFormatStructureDefinition.NUMBERINGCAPTIONSTRING, abstNum);
				
				abstNum = Context.getWmlObjectFactory().createNumberingAbstractNum();
				abstNum.setAbstractNumId(BigInteger.valueOf(96));
				abstNum.setMultiLevelType(Context.getWmlObjectFactory().createNumberingAbstractNumMultiLevelType());
				abstNum.getMultiLevelType().setVal("singleLevel");
				styleMap.getNumberingMap().put(AcademicFormatStructureDefinition.NUMBERINGFIGURESTRING, abstNum);
				for (Object key : jsonRootObject.keySet()) {
					Style style = Context.getWmlObjectFactory().createStyle();
					style.setPPr(Context.getWmlObjectFactory().createPPr());
//					style.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
					style.setRPr(Context.getWmlObjectFactory().createRPr());
					String keyStr = (String) key;
					Object value = jsonRootObject.get(keyStr);
					if(value instanceof JSONObject){
						JSONObject jsonObjItem = (JSONObject) value;
						loadDetails(styleMap, jsonObjItem, style, keyStr);
					} else if(value instanceof String){
						if(((String) key).equals("stylename")){
							styles.put((String) value, styleMap);
						} else {
							
						}
						
					}
				}
				/*
				 * do some intelligent adapt 
				 */
				autoFillAndModify(styleMap);
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
	
	public void loadDefinition(){
		try {
//			File styleFile = new File(getClass().getResource("/sources/acmStyleDefinition.json").toURI());
//			loadFile(styleFile);
//			String filePath = this.getClass().getClassLoader().getResource("/sources/acmStyleDefinitionold.json").getFile();
			InputStream in = this.getClass().getResourceAsStream("/sources/acmStyleDefinitionold.json");
			File styleFile = File.createTempFile("acmStyleDefinitionold", ".tmp");
			FileOutputStream out = null;
			try {
				out = new FileOutputStream(styleFile);
				IOUtils.copy(in, out);
	        } finally {
	        	if(in!=null){
	        		in.close();
	        	}
	        	if(out!=null){
	        		out.close();
	        	}
			}
			loadFile(styleFile);
			styleFile.delete();
			
//			styleFile = new File(this.getClass().getResource("/sources/ieeeStyleDefinition.json").toURI());
			in = this.getClass().getResourceAsStream("/sources/ieeeStyleDefinition.json");
			styleFile = File.createTempFile("ieeeStyleDefinition", ".tmp");
			out = null;
			try {
				out = new FileOutputStream(styleFile);
				IOUtils.copy(in, out);
	        } finally {
	        	if(in!=null){
	        		in.close();
	        	}
	        	if(out!=null){
	        		out.close();
	        	}
			}
			loadFile(styleFile);
			styleFile.delete();
			
//			styleFile = new File(this.getClass().getResource("/sources/springerStyleDefinition.json").toURI());
			in = this.getClass().getResourceAsStream("/sources/springerStyleDefinition.json");
			styleFile = File.createTempFile("springerStyleDefinition", ".tmp");
			out = null;
			try {
				out = new FileOutputStream(styleFile);
				IOUtils.copy(in, out);
	        } finally {
	        	if(in!=null){
	        		in.close();
	        	}
	        	if(out!=null){
	        		out.close();
	        	}
			}
			loadFile(styleFile);
			styleFile.delete();
			
//			styleFile = new File(this.getClass().getResource("/sources/elsevierStyleDefinition.json").toURI());
			in = this.getClass().getResourceAsStream("/sources/elsevierStyleDefinition.json");
			styleFile = File.createTempFile("elsevierStyleDefinition", ".tmp");
			out = null;
			try {
				out = new FileOutputStream(styleFile);
				IOUtils.copy(in, out);
	        } finally {
	        	if(in!=null){
	        		in.close();
	        	}
	        	if(out!=null){
	        		out.close();
	        	}
			}
			loadFile(styleFile);
			styleFile.delete();
			
			styleFile = AppEnvironment.getInstance().getNewStyleFile();
			if(styleFile!=null){
				loadFile(styleFile);
				LogUtils.log("new style file is loaded.");
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
		
	}
	
	public void loadCDefinition(){
		try {
			File styleFile = AppEnvironment.getInstance().getNewStyleFile();
			if (styleFile != null) {
				loadFile(styleFile);
				LogUtils.log("new style file is loaded.");
			}
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
	}
}
