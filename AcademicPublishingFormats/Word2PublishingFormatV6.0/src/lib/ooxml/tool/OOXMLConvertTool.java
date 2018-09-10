package lib.ooxml.tool;

import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.PPrBase.NumPr.Ilvl;
import org.docx4j.wml.PPrBase.NumPr.NumId;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.SectPr.Type;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.PPrBase.TextAlignment;
import org.docx4j.wml.ParaRPr;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import tools.LogUtils;

import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RStyle;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.STHeightRule;
import org.docx4j.wml.STWrap;
import org.docx4j.wml.STXAlign;
import org.docx4j.wml.STYAlign;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.U;
import org.docx4j.wml.UnderlineEnumeration;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.bind.JAXBElement;
import org.docx4j.TraversalUtil;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.ClassFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTColumns;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTSimpleField;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Jc;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.Num;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.R.LastRenderedPageBreak;
import org.docx4j.wml.SdtPr.Bibliography;

public class OOXMLConvertTool {

	public static String sourceFormat = "";
	
	public final static String SOURCEFORMAT_ACM = "acm";
	public final static String SOURCEFORMAT_ELSEVIER = "elsevier";
	public final static String SOURCEFORMAT_IEEE = "ieee";
	public final static String SOURCEFORMAT_SPRINGER = "springer";
	
	public static boolean isStyleExist(Style styleItem, StyleDefinitionsPart stylePart){
		boolean isExist = false;
		
		if(stylePart!=null){
			Style style = stylePart.getStyleById(styleItem.getStyleId());
			if(style!=null){
				isExist = true;
			}
		}
		
		return isExist;
	}
	
	public static void resetStyle(Style style) {
		style.setPPr(null);
		style.setRPr(null);
	}
	
	public static PStyle getPStyle(PPr pPr, boolean isForcedInstance) {
		PStyle pStyle = null;
		
		pStyle = pPr.getPStyle();
		if(pStyle==null && isForcedInstance) {
			pStyle = Context.getWmlObjectFactory().createPPrBasePStyle();
			pPr.setPStyle(pStyle);
		}
		
		return pStyle;
	}
	
	public static PPr getPPr(P p, boolean isForcedInstance, boolean isForcedReset){
		PPr pPr = null;
		
		if(isForcedReset){
			p.setPPr(Context.getWmlObjectFactory().createPPr());
		}
		
		pPr = p.getPPr();
		if(pPr==null && isForcedInstance){
			pPr = Context.getWmlObjectFactory().createPPr();
			p.setPPr(pPr);
		}
		
		return pPr;
	}
	
	public static boolean removeCaptionNum(P p){
		boolean result = false;
		try {
			String regex = "^Fig.\\s{1,1}\\d+.";
			boolean isBold = true;
			boolean isItalic = false;
			result = removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
			if(!result){
				regex = "Figure\\s{1,1}\\d.";
				result = removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
			}
			
			if(!result){
				regex = "Table\\s{1,1}\\d.";
				result = removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
			}
			
			if(!result){
				regex = "TABLE\\s{1,1}\\d.";
				result = removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
			}
			
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return result;
	}
	
	public static boolean removeHeadingNum(int headingLvl, P p){
		boolean result = false;
		if(p!=null){
			boolean isBold = false;
			boolean isItalic = false;
			String regex = "";
			try {
				switch (headingLvl) {
				case 1:
					// ACM
					isBold = true;
					isItalic = false;
					regex = "^\\d+.";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Elsevier
					isBold = true;
					isItalic = false;
					regex = "^\\d+.";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// IEEE
					isBold = false;
					isItalic = false;
					regex = "^\\d";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Springer
					isBold = true;
					isItalic = false;
					regex = "^\\d";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					break;
				case 2:
					// ACM
					isBold = true;
					isItalic = false;
					regex = "^\\d+.\\d+";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Elsevier
					isBold = false;
					isItalic = true;
					regex = "^\\d+.\\d+.";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// IEEE
					isBold = true;
					isItalic = false;
					regex = "^[A-Z].{1,1}?";
					OOXMLConvertTool.removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Springer
					isBold = true;
					isItalic = false;
					regex = "^\\d+.\\d+";
					OOXMLConvertTool.removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					break;
				case 3:
					// ACM
					isBold = false;
					isItalic = true;
					regex = "^\\d+.\\d+.\\d+";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Elsevier
					// IEEE
					isBold = false;
					isItalic = true;
					regex = "^[a-z]\\){1,1}?";
					OOXMLConvertTool.removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Springer
					break;
				case 4:
					// ACM
					isBold = false;
					isItalic = true;
					regex = "^\\d+.\\d+.\\d+.\\d+";
					removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Elsevier
					// IEEE
					isBold = false;
					isItalic = true;
					regex = "^\\d\\){1,1}?";
					OOXMLConvertTool.removeSectionNumStringOfHeading(p, regex, isBold, isItalic);
					// Springer
					break;
				default:
					break;
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
			result = true;
		}
		return result;
	}
	
	public static void removeAllRPROfP(P p){
		List<Object> pContents = p.getContent();
		for(Object o : pContents){
			if(o instanceof R){
				R r = (R)o;
				r.setRPr(null);
			}
		}
	}
	
	public static boolean removeSectionFromText(P p, StyleDefinitionsPart stylePart, String regex, boolean isBold, boolean isItalic){
		boolean result = false;

		if(p!=null && p.getContent()!=null){
			int indexOfHeadLastNode = getIndexOfLastSameRStyle(p, stylePart, isBold, isItalic);
			if(indexOfHeadLastNode>=0){
				for (int i = 0; i < indexOfHeadLastNode; i++) {
					p.getContent().remove(0);
				}
				result = true;
//				P newP = Context.getWmlObjectFactory().createP();
//				R newR = Context.getWmlObjectFactory().createR();
//				newP.setPPr(Context.getWmlObjectFactory().createPPr());
//				if(headString.matches(regex+".*$")){
//					headString.replaceAll(regex, "");
//					text.setValue(headString);
//					result = true;
//				} else if(headString.matches("^\\d\\)\\s.*$")){
//					headString.replaceAll("^\\d\\)\\s", "");
//					text.setValue(headString);
//					result = true;
//				}
			}
		}
		
		return result;
	}
	public static String removeSectionNumStringOfHeading(String str, String regex){
		String result = str;
		
		if(str.matches(regex+".*$")){
			result = str.replaceAll(regex, "");
		} else if(str.matches("^\\d\\)\\s.*$")){
			result = str.replaceAll("^\\d\\)\\s", "");
		}

		return result;
	}
	public static boolean removeSectionNumStringOfHeading(P p, String regex, boolean isBold, boolean isItalic){
		boolean result = false;
		
		int i = 0;
		
		if(p.getContent()!=null){
			while(i<p.getContent().size()){
				Object o = p.getContent().get(i);
				if(o instanceof R){
					R r = (R)o;
					List<Object> rContents = r.getContent();
					for(Object ro : rContents){
						if(ro instanceof JAXBElement<?>){
							Object deepO = ((JAXBElement<?>)ro).getValue();
							if(deepO instanceof Text){
								Text text = (Text) deepO;
								String str = text.getValue();
								if(str.matches(regex+".*$")){
									str.replaceAll(regex, "");
									text.setValue(str);
									result = true;
								} else if(str.matches("^\\d\\)\\s.*$")){
									str.replaceAll("^\\d\\)\\s", "");
									text.setValue(str);
									result = true;
								}
							}
						}
					}
				}
				i++;
			}
		}
		
		return result;
	}
	
	public static P getNextP(MainDocumentPart documentPart, int docContentIndex){
		P p = null;
		List<Object> docContent = documentPart.getContent();
		while(++docContentIndex<docContent.size()){
			Object o = docContent.get(docContentIndex);
			if(o instanceof P){
				p = (P)o;
				break;
			}
		}
		return p;
	}
	public static boolean insertPAt(MainDocumentPart documentPart, P p, int docContentIndex){
		boolean isInserted = false;
		if(p!=null){
			documentPart.getContent().add(docContentIndex, p);
			isInserted = true;
		}
		return isInserted;
	}
	public static boolean insertRIntoP(P p, R r, int index){
		boolean isInserted = false;
		
		if(p.getContent()!=null && p!=null && r!=null){
			p.getContent().add(index, r);
			isInserted = true;
		}
		
		return isInserted;
	}
	
	public static boolean removePAt(MainDocumentPart documentPart, int docContentIndex){
		boolean isRemoved = false;
		
		try {
			documentPart.getContent().remove(docContentIndex);
			isRemoved = true;
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return isRemoved;
	}
	private static boolean chkRStyleInR(R r, Style chkStyle, StyleDefinitionsPart stylePart){
		boolean isAccepted = true;
		RPr rPr = null;
		if(r!=null && r.getRPr()!=null){
			rPr = r.getRPr();
		}
		if(rPr != null && rPr.getRStyle()!=null && chkStyle!=null && chkStyle.getRPr()!=null){
			Style style = stylePart.getStyleById(rPr.getRStyle().getVal());
			if(chkStyle.getRPr().getB()!=null){
				if(chkStyle.getRPr().getB().equals(style.getRPr().getB())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getI()!=null){
				if(chkStyle.getRPr().getI().equals(style.getRPr().getI())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getSz()!=null && chkStyle.getRPr().getSz().getVal()!=null){
				if(chkStyle.getRPr().getSz().equals(style.getRPr().getSz())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getRFonts()!=null && chkStyle.getRPr().getRFonts().getAscii()!=null){
				if(chkStyle.getRPr().getRFonts().getAscii().equals(style.getRPr().getRFonts().getAscii())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getCaps()!=null){
				if(chkStyle.getRPr().getCaps().isVal() && style.getRPr().getCaps()!=null && style.getRPr().getCaps().isVal()){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getSmallCaps()!=null){
				if(chkStyle.getRPr().getSmallCaps().isVal() && style.getRPr().getSmallCaps()!=null && style.getRPr().getSmallCaps().isVal()){
					
				} else {
					isAccepted = false;
				}
			}
		} else if(r.getRPr()!=null) {
			if(chkStyle.getRPr().getB()!=null){
				if(chkStyle.getRPr().getB().equals(r.getRPr().getB())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getI()!=null){
				if(chkStyle.getRPr().getI().equals(r.getRPr().getI())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getSz()!=null && chkStyle.getRPr().getSz().getVal()!=null){
				if(r.getRPr().getSz()!=null && r.getRPr().getSz().getVal()!=null 
						&& chkStyle.getRPr().getSz().getVal().intValue()==r.getRPr().getSz().getVal().intValue()){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getRFonts()!=null && chkStyle.getRPr().getRFonts().getAscii()!=null){
				if(chkStyle.getRPr().getRFonts().getAscii().equals(r.getRPr().getRFonts().getAscii())){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getCaps()!=null){
				if(chkStyle.getRPr().getCaps().isVal() && r.getRPr().getCaps()!=null && r.getRPr().getCaps().isVal()){
					
				} else {
					isAccepted = false;
				}
			}
			if(chkStyle.getRPr().getSmallCaps()!=null){
				if(chkStyle.getRPr().getSmallCaps().isVal() && r.getRPr().getSmallCaps()!=null && r.getRPr().getSmallCaps().isVal()){
					
				} else {
					isAccepted = false;
				}
			}
		} else {
			isAccepted = false;
		}
		
		return isAccepted;
	}
	private static boolean isHasStyleInDeep(P p, Style chkStyle, StyleDefinitionsPart stylePart){
		boolean isAccepted = false;
		if(p!=null && p.getContent()!=null && p.getContent().size()>0 && chkStyle!=null){
			Style style = null;
			if(chkStyle.getStyleId()!=null){
				style = stylePart.getStyleById(chkStyle.getStyleId());
			}
			for (int i = 0; i < p.getContent().size(); i++) {
				Object o = p.getContent().get(i);
				if(o instanceof R){
					R r = (R)o;
					String rText = getRText(r);
					if(rText!=null && rText.length()>0){
						
					} else {
						continue;
					}
					if(r.getRPr()!=null && chkStyle.getRPr()!=null){
						RPr rPr = r.getRPr();
						RPr chkRPr = chkStyle.getRPr();
						RPr styleRPr = style!=null?style.getRPr():null;
						if(chkRPr.getB()!=null){
							if(chkRPr.getB().isVal() && rPr.getB()!=null && rPr.getB().isVal()){
								isAccepted = true;
								break;
							} else if(styleRPr!=null && chkRPr.getB().isVal() && styleRPr.getB()!=null && styleRPr.getB().isVal()){
								isAccepted = true;
								break;
							} else {
								break;
							}
						}
						if(chkRPr.getI()!=null){
							if(chkRPr.getI().isVal() && rPr.getI()!=null && rPr.getI().isVal()){
								isAccepted = true;
								break;
							} else if(styleRPr!=null && chkRPr.getI().isVal() && styleRPr.getI()!=null && styleRPr.getI().isVal()){
								isAccepted = true;
								break;
							} else {
								break;
							}
						}
						if(chkRPr.getRFonts()!=null && chkRPr.getRFonts().getAscii()!=null){
							if(rPr.getRFonts()!=null
									&& rPr.getRFonts().getAscii()!=null && chkRPr.getRFonts().getAscii().equals(rPr.getRFonts().getAscii())){
								isAccepted = true;
								break;
							} else if(styleRPr!=null && styleRPr.getRFonts()!=null
									&& styleRPr.getRFonts().getAscii()!=null && chkRPr.getRFonts().getAscii().equals(styleRPr.getRFonts().getAscii())){
								isAccepted = true;
								break;
							} else {
								break;
							}
						}
						if(chkRPr.getSz()!=null && rPr.getSz()!=null
								&& chkRPr.getSz().getVal()!=null && rPr.getSz().getVal()!=null){
							if(chkRPr.getSz().getVal().intValue() == rPr.getSz().getVal().intValue()){
								isAccepted = true;
								break;
							} else if(styleRPr!=null && styleRPr.getSz()!=null && styleRPr.getSz().getVal()!=null &&
									chkRPr.getSz().getVal().intValue() == styleRPr.getSz().getVal().intValue()){
								isAccepted = true;
								break;
							} else {
								break;
							}
						}
						if(chkRPr.getSmallCaps()!=null){
							if(rPr.getSmallCaps()!=null && chkRPr.getSmallCaps().isVal() && rPr.getSmallCaps().isVal()){
								isAccepted = true;
								break;
							} else if(styleRPr!=null && styleRPr.getSmallCaps()!=null && chkRPr.getSmallCaps().isVal() && styleRPr.getSmallCaps().isVal()){
								isAccepted = true;
								break;
							} else {
								break;
							}
						}
						if(chkRPr.getCaps()!=null){
							if(rPr.getCaps()!=null && chkRPr.getCaps().isVal() && rPr.getCaps().isVal()){
								isAccepted = true;
								break;
							} else if(styleRPr!=null && styleRPr.getCaps()!=null && chkRPr.getCaps().isVal() && styleRPr.getCaps().isVal()){
								isAccepted = true;
								break;
							} else {
								break;
							}
						}
					}
				}
			}
		}
		return isAccepted;
	}
	private static boolean chkStyleInR(P p, Style chkStyle, StyleDefinitionsPart stylePart){
		boolean isAccepted = true;
		
		if(p!=null && p.getContent()!=null && p.getContent().size()>0 && chkStyle!=null){
			for (int i = 0; i < p.getContent().size(); i++) {
				Object o = p.getContent().get(i);
				if(o instanceof R){
					R r = (R)o;
					if(r.getRPr()!=null && chkStyle.getRPr()!=null){
						RPr rPr = r.getRPr();
						RPr chkRPr = chkStyle.getRPr();
						if(chkRPr.getB()!=null){
							if(chkRPr.getB().isVal() && rPr.getB()!=null && rPr.getB().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(chkRPr.getI()!=null){
							if(chkRPr.getI().isVal() && rPr.getI()!=null && rPr.getI().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(chkRPr.getRFonts()!=null && chkRPr.getRFonts().getAscii()!=null){
							if(rPr.getRFonts()!=null
									&& rPr.getRFonts().getAscii()!=null && chkRPr.getRFonts().getAscii().equals(rPr.getRFonts().getAscii())){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(chkRPr.getSz()!=null && rPr.getSz()!=null
								&& chkRPr.getSz().getVal()!=null && rPr.getSz().getVal()!=null){
							if(chkRPr.getSz().getVal().intValue() == rPr.getSz().getVal().intValue()){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(chkRPr.getSmallCaps()!=null && rPr.getSmallCaps()!=null){
							if(chkRPr.getSmallCaps().isVal() && rPr.getSmallCaps().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(chkRPr.getCaps()!=null && rPr.getCaps()!=null){
							if(chkRPr.getCaps().isVal() && rPr.getCaps().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(chkRPr.getB()!=null || chkRPr.getI()!=null 
								|| chkRPr.getSz()!=null || chkRPr.getRFonts()!=null){
							isAccepted = false;
						} else {
							isAccepted = chkRStyleInR(r, chkStyle, stylePart);
						}
						break;
					} else {
						isAccepted = false;
						break;
					}
				}
			}
		} else {
			isAccepted = false;
		}
		
		return isAccepted;
	}
	/*
	public static UserMessage modifyStyleDoc(Style styleItem, StyleDefinitionsPart stylePart) {
		UserMessage msg = new UserMessage();
		
		RPr rPr = null;
		PPr pPr = null;
		if(isStyleExist(styleItem, stylePart)){
			Style style = stylePart.getStyleById(styleItem.getStyleId());
			rPr = getRPr(style);
			pPr = getPPr(style);
			
		} else {
			msg.setMessageCode(UserMessage.MESSAGE_UNABLE_FOUND_STYLE_BY_CLZID);
			return msg;
		}
		
		msg.setMessageDetails("OK, style is modified.");
		return msg;
		
	}*/
	public static void createStyle(Style styleItem, StyleDefinitionsPart stylePart) {
		if(styleItem!=null){
			if(styleItem.getStyleId()!=null && styleItem.getStyleId().length()>0){
				try{
//					documentPart.addStyledParagraphOfText(styleItem.getStyleId(), "");
					
					if(stylePart.getContents()!=null){
						stylePart.getContents().getStyle().add(styleItem);
					}
//					modifyFont(stylePart, itemStyle);
				} catch	(Exception e) {
					LogUtils.log("cannot add new style item " + e.getMessage());
					UserMessage msg = new UserMessage();
					msg.setMessageCode(UserMessage.MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM);
				}
			} else {
				UserMessage msg = new UserMessage();
				msg.setMessageCode(UserMessage.MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM_NONE_CID);
			}
		} else {
			UserMessage msg = new UserMessage();
			msg.setMessageCode(UserMessage.MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM_NULL);
		}
	}
	private static PPr getPPr(Style style){
		PPr pPr = null;
		if(style.getPPr()!=null){
			pPr = style.getPPr();
		} else {
			pPr = new PPr();
			style.setPPr(pPr);
		}
		return pPr;
	}
	
	private static RPr getRPr(Style style){
		RPr rPr = null;
		if(style!=null){
			if(style.getRPr()!=null){
				rPr = style.getRPr();
			} else {
				rPr = new RPr();
				style.setRPr(rPr);
			}
		}
		return rPr;
	}

//	public static RPr getRPr(String num){
//		RPr rPr = null;
//		
//		
//		return rPr;
//	}
	
	public static String getSectionHeader3String(P p, StyleDefinitionsPart stylePart){
		String str = "";
		
		// IEEE
		str = getSectionHeaderStr(p, stylePart, false, true);
		// Springer
		if(str!=null && str.length()==0 || str == null){
			str = getSectionHeaderStr(p, stylePart, true, false);
		}
		
		return str;
	}
	
	public static String getSectionHeader4String(P p, StyleDefinitionsPart stylePart){
		String str = "";
		
		// IEEE
		str = getSectionHeaderStr(p, stylePart, false, true);
		if(str!=null && str.length()==0 || str == null){
			str = getSectionHeaderStr(p, stylePart, false, true);
		}
		
		return str;
	}
	
	private static String getTextWithSpecialStyle(P p, StyleDefinitionsPart stylePart, Style specialStyle){
		String result = "";
		if(p.getPPr()!=null){
			if(specialStyle!=null && specialStyle.getRPr()!=null){
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null && r.getRPr().getB()!=null && r.getRPr().getB().isVal()
								&& specialStyle.getRPr().getB()!=null && specialStyle.getRPr().getB().isVal()
								){
							result += getRText(r);
						} else if(r.getRPr()!=null && r.getRPr().getI()!=null && r.getRPr().getI().isVal()
								&& specialStyle.getRPr().getI()!=null && specialStyle.getRPr().getI().isVal()) {
							result += getRText(r);
						}
					}
				}
			}
		}
		return result;
	}
	private static boolean isIncludeSEQ(P p){
		boolean result = false;
		
		if(p!=null){
			List<Object> pContent = p.getContent();
			if(pContent!=null){
				for(Object pO : pContent){
					if(pO instanceof R){
						R r = (R)pO;
						if(r.getContent()!=null){
							for(Object rO : r.getContent()){
								if(rO instanceof JAXBElement<?>){
									if(((JAXBElement<?>)rO).getValue() instanceof Text && ((JAXBElement<?>)rO).getName().getLocalPart().equals("instrText")){
										Text text = (Text) ((JAXBElement<?>)rO).getValue();
										if(text.getValue().indexOf("SEQ")>=0){
											result = true;
										}
									} else if(((JAXBElement<?>)rO).getValue() instanceof Text){
										// normal text ignore
									} else if(((JAXBElement<?>)rO).getValue() instanceof FldChar){
//										FldChar fldChar = (FldChar) ((JAXBElement<?>)rO).getValue();
//										System.out.println(fldChar.getFldCharType().name());
										// mark the field begin and end
									}
								} else {
//									System.out.println(rO.getClass().getName());
								}
							}
						}
					} else {
//						System.out.println(pO.getClass().getName());
					}
				}
			}
		}
		
		return result;
	}
	public static String getPText(P p){
		String result = "";
		if(p!=null){
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						R r = (R)o;
						result += getRText(r);
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R sfR = (R)sfO;
										result+=getRText(sfR);
									}
								}
							}
							
						}
					} else {
//						System.out.println(o.getClass().getName());
					}
				}
				
			}
		}
		return result;
	}
	public static String getRText(R r){
		String result = "";
		if(r!=null){
			List<Object> rContents = r.getContent();
			if(rContents!=null && rContents.size()>0){
				for(Object o : rContents){
					if(o instanceof JAXBElement<?>){
						JAXBElement<?> e = (JAXBElement<?>)o;
						if(e.getValue() instanceof Text && !e.getName().getLocalPart().equals("instrText")){
							Text text = (Text) e.getValue();
							result += text.getValue();
						}
					} else {
						if(o instanceof Text){
							Text text = (Text) o;
							result += text.getValue();
						}
					}
				}
			}
		}
		return result;
	}
	public static boolean removeUselessBreak(P p){
		boolean result = false;
				
		if( p != null && p.getContent()!=null){
			for (int i = 0; i < p.getContent().size(); i++) {
				Object o = p.getContent().get(i);
				if(o instanceof R){
					R r = (R)o;
					if(r.getContent()!=null && r.getContent().size()>0){
						for (int j = 0; j < r.getContent().size(); j++) {
							Object oR = r.getContent().get(j);
							if(oR instanceof JAXBElement<?>){
								if(((JAXBElement) oR).getValue() instanceof LastRenderedPageBreak){
									r.getContent().remove(j);
									j--;
								}
							} else if(oR instanceof Br){
								r.getContent().remove(j);
								j--;
							}
						}
					}
				}
			}
		}
		
		return result;		
	}
	public static boolean appendBr(P p, boolean isColumn){
		boolean result = false;
		
		if(p!=null){
			for(Object o : p.getContent()){
				if(o instanceof R){
					R r = (R)o;
					Br br = Context.getWmlObjectFactory().createBr();
					if(isColumn) {
						br.setType(STBrType.COLUMN);
					} else {
						br.setType(STBrType.PAGE);
					}
					LastRenderedPageBreak lrpb = Context.getWmlObjectFactory().createRLastRenderedPageBreak();
					r.getContent().add(lrpb);
					r.getContent().add(br);
				}
			}
		}
		
		return result;
	}
	private static boolean isREmpty(R r){
		boolean result = false;
		
		if(r!=null){
			String rText = getRText(r);
			if(rText!=null && rText.length()>0){
				
			} else {
				result = true;
			}
		}
		
		return result;
	}
	
	public static int getIndexOfLastSameRStyle(P p, StyleDefinitionsPart stylePart, boolean isB, boolean isI){
		int rIndex = -1;
		
		if(p!=null){
			List<Object> pContents = p.getContent();
			Style style = null;
			RStyle rStyle = null;
			int i=0;
			
			if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
				style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
			}
			
			
			for(;i<pContents.size();i++){
				Object o = pContents.get(i);
				if(o instanceof R){
					R r = (R)o;
					if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
						rStyle = r.getRPr().getRStyle();
					}
					if(isREmpty(r)){
						continue;
					}
					if(style != null){
						if(r.getRPr()!=null){
							if(style.getRPr()!=null
								&& r.getRPr().getB()!=null && r.getRPr().getB().isVal() && style.getRPr().getB().isVal() 
								&& isB){
								// same, ignore
							} else if(style.getRPr()!=null
									&& r.getRPr().getI()!=null && r.getRPr().getI().isVal()
									&& style.getRPr().getI().isVal()
									&& isI){
								// same, ignore	
							} else {
								rIndex = i;
								break;
							}
						} else if(rStyle != null){
							Style rStyleDef = stylePart.getStyleById(rStyle.getVal());
							if(r.getRPr()!=null && r.getRPr().getRStyle()!=null 
									&& rStyle.getVal().equals(r.getRPr().getRStyle().getVal())){
								// same, ignore
							} else if(rStyleDef!=null){
								if(style.getRPr()!=null && style.getRPr().getB()!=null && style.getRPr().getB().isVal() 
										&& rStyleDef.getRPr()!=null && rStyleDef.getRPr().getB()!=null && rStyleDef.getRPr().getB().isVal() 
										&& isB){
									// same, ignore
								} else if(style.getRPr()!=null && style.getRPr().getI()!=null &&style.getRPr().getI().isVal()
										&& rStyleDef.getRPr()!=null && rStyleDef.getRPr().getB()!=null && rStyleDef.getRPr().getB().isVal()
										&& isI) {
									// same, ignore
								} else {
									rIndex = i;
									break;
								}
							}
						} else {
							rIndex = i>0?i:(i+1);
							break;
						}
						
					} else if(style == null && rStyle == null){
						if(r.getRPr()!=null){
							if(r.getRPr().getB()!=null && r.getRPr().getB().isVal() && isB){
								if(style == null){
									style = new Style();
									style.setRPr(new RPr());
								}
								style.getRPr().setB(new BooleanDefaultTrue());
							}
							if(r.getRPr().getI()!=null && r.getRPr().getI().isVal() && isI){
								if(style == null){
									style = new Style();
									style.setRPr(new RPr());
								}
								style.getRPr().setI(new BooleanDefaultTrue());
							}
							if(r.getRPr().getRStyle()!=null){
								rStyle = r.getRPr().getRStyle();
							}
							if(style != null || r.getRPr().getRStyle()!=null){
							}
						}
					} else {
						rIndex = i;
						break;
					}
				}
			}
			if(style == null && rStyle == null){
				
			} else {
				if(rIndex<=0){
					rIndex = i;
				} else {
					
				}
			}
		}
		
		return rIndex;
	}
	
	public static int getIndexOfLastSameRStyle(P p, StyleDefinitionsPart stylePart, Style foundStyle){
		int rIndex = -1;
		
		if(p!=null){
			List<Object> pContents = p.getContent();
			
//			Style style = null;
			RStyle rStyle = null;
			int i=0;
			
			for(;i<pContents.size();i++){
				Object o = pContents.get(i);
				if(o instanceof R){
					R r = (R)o;
					if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
						rStyle = r.getRPr().getRStyle();
					}
					if(isREmpty(r)){
						continue;
					}
					if(r.getRPr()!=null){
						if(r.getRPr().getB()!=null && r.getRPr().getB().isVal() && foundStyle.getRPr().getB()!=null && foundStyle.getRPr().getB().isVal()){
							
						} else if(r.getRPr().getB()==null && foundStyle.getRPr().getB()==null){
							
						} else {
							rIndex = i;
							break;
						}
						if(r.getRPr().getI()!=null && r.getRPr().getI().isVal() && foundStyle.getRPr().getI()!=null && foundStyle.getRPr().getI().isVal()){
							
						} else if(r.getRPr().getI()==null && foundStyle.getRPr().getI()==null){
							
						} else {
							rIndex = i;
							break;
						}
					} else if(rStyle != null){
						Style rStyleDef = stylePart.getStyleById(rStyle.getVal());
						if(r.getRPr()!=null && r.getRPr().getRStyle()!=null 
								&& rStyle.getVal().equals(r.getRPr().getRStyle().getVal())){
							// same, ignore
						} else if(rStyleDef!=null){
							if(foundStyle.getRPr()!=null && foundStyle.getRPr().getB()!=null && foundStyle.getRPr().getB().isVal() 
									&& rStyleDef.getRPr()!=null && rStyleDef.getRPr().getB()!=null && rStyleDef.getRPr().getB().isVal() 
									){
								// same, ignore
							} else if(foundStyle.getRPr()!=null && foundStyle.getRPr().getI()!=null &&foundStyle.getRPr().getI().isVal()
									&& rStyleDef.getRPr()!=null && rStyleDef.getRPr().getB()!=null && rStyleDef.getRPr().getB().isVal()
									) {
								// same, ignore
							} else {
								rIndex = i;
								break;
							}
						}
					}
				}
			}
			if(rStyle == null){
				
			} else {
				if(rIndex<=0){
					rIndex = i;
				} else {
					
				}
			}
		}
		
		return rIndex;
	}
	
	public static String getSectionHeaderStr(P p, StyleDefinitionsPart stylePart, boolean isB, boolean isI){
		String str = "";
		
		if(p!=null){
			List<Object> pContents = p.getContent();
			int rIndex = getIndexOfLastSameRStyle(p, stylePart, isB, isI);
//			rIndex = Math.min(rIndex, pContents.size()-1);
			for(int i=0; i<rIndex; i++){
				Object o = pContents.get(i);
				if(o instanceof R){
					R r = (R)o;
					List<Object> rContents = r.getContent();
					for(Object ro : rContents){
						if(ro instanceof JAXBElement<?>){
							Object deepO = ((JAXBElement<?>)ro).getValue();
							if(deepO instanceof Text){
								Text text = (Text) deepO;
								str += text.getValue();
							}
						}
					}
				}
			}
		}
		
		return str;
	}
	
	public static String getSectionHederStr(P p, StyleDefinitionsPart stylePart, Style style){
		String str = "";
		
		if(p!=null){
			List<Object> pContents = p.getContent();
			int rIndex = getIndexOfLastSameRStyle(p, stylePart, style);
			for(int i=0; i<rIndex; i++){
				Object o = pContents.get(i);
				if(o instanceof R){
					R r = (R)o;
					List<Object> rContents = r.getContent();
					for(Object ro : rContents){
						if(ro instanceof JAXBElement<?>){
							Object deepO = ((JAXBElement<?>)ro).getValue();
							if(deepO instanceof Text){
								Text text = (Text) deepO;
								str += text.getValue();
							}
						}
					}
				}
			}
		}
		
		return str;
	}
	
	private static boolean isPStyleWithSectionNum(NumberingDefinitionsPart numberingPart, BigInteger numId, BigInteger iLvl, String pattern, NumberFormat numFm){
		boolean isAccepted = false;
		try {
			Numbering numbering = numberingPart.getContents();
			if(numbering!=null){
				List<Num> numList = numbering.getNum();
				BigInteger abstractNumId = null;
				for(int numIndex=0; numIndex<numList.size(); numIndex++) {
					if(numList.get(numIndex).getNumId().intValue() == numId.intValue()) {
						abstractNumId = numList.get(numIndex).getAbstractNumId().getVal();
						break;
					}
				}
				if(abstractNumId!=null){
//					HashMap<String, AbstractListNumberingDefinition> ndpMap = numberingPart.getAbstractListDefinitions();
//					if(ndpMap!=null) {
//						ndpMap.get(String.valueOf(abstractNumId));
//						for (Map.Entry<String, AbstractListNumberingDefinition> entry : ndpMap.entrySet()) {
//							
//						}
//					}
					List<AbstractNum> abstractNumList = numbering.getAbstractNum();
					for(int aNumIndex=0; aNumIndex<abstractNumList.size(); aNumIndex++) {
						if(abstractNumList.get(aNumIndex).getAbstractNumId().intValue() == abstractNumId.intValue()){
							List<Lvl> lvlList = abstractNumList.get(aNumIndex).getLvl();
							for(int lvlIndex=0; lvlIndex<lvlList.size(); lvlIndex++) {
								if(lvlList.get(lvlIndex).getIlvl().intValue()==iLvl.intValue()){
									if(lvlList.get(lvlIndex).getLvlText().getVal().equals(pattern) 
											&& lvlList.get(lvlIndex).getNumFmt().getVal().equals(numFm)){
										isAccepted = true;
										break;
									}
								}
							}
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return isAccepted;
	}
	
	private static String getSectionCaptialNum(P p){
		String str = "";
		if(p!=null){
			if(p.getContent()!=null && p.getContent().size()>0){
				String s = p.toString();
				int pos = 0;
				while(pos<s.length()){
					String ts = s.substring(pos, ++pos);
					try {
						if(ts.equals(".")||ts.equals(")")){
							// ignore
						} else if(ts.matches("^(M|D|C|L|X|V|I)")){
							// ignore
						} else if(Integer.parseInt(ts)>=0){
							// ignore
						} else if(ts.equals(" ")){
							break;
						}
						str += ts;
					} catch (Exception e) {
						// TODO: handle exception
						// not integer
						str = "";
						break;
					}
				}
			}
		}
		return str;
	}
	
	public static boolean isSectionNum(String testString, String pattern){
		boolean isAccepted = false;
		
		String regex = "^\\d+(\\.\\d+)*?";
		String regex1 = "^[a-z][\\)]{1,1}?";
		String regex2 = "^[A-Z][\\)]{1,1}?";
		String regex3 = "^\\d+[\\)]{1,1}?";
		String regex4 = "^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$";
		
		if(testString.endsWith(".")){
			testString = testString.substring(0, testString.length()-1);
		}
		if(testString!=null && testString.length()==0){
			isAccepted = false;
		} else if(testString.matches(regex)){
			isAccepted = true;
		} else if(testString.matches(regex1)){
			isAccepted = true;
		} else if(testString.matches(regex2)){
			isAccepted = true;
		} else if(testString.matches(regex3)){
			isAccepted = true;
		} else if(testString.matches(regex4)){
			isAccepted = true;
		} else if(pattern!=null && testString.matches(pattern)){
			isAccepted = true;
		} else {
		}
		
		return isAccepted;
	}
	
	public static boolean isPTextStartWithString(P p, String str, boolean isStrictly){
		boolean isAccepted = false;
		
		if(p!=null && p.getContent()!=null && p.getContent().size()>0){
			if(isStrictly){
				if(p.toString().indexOf(str)==0){
					isAccepted = true;
				}
			} else {
				if(p.toString().toUpperCase().indexOf(str.toUpperCase())==0){
					isAccepted = true;
				}
			}
		}
		
		return isAccepted;
	}
	
	public static boolean isPTextIncludeString(P p, String str, boolean isStrictly) {
		boolean isAccepted = false;
		
		if(p!=null && p.getContent()!=null && p.getContent().size()>0) {
			if(isStrictly){
				if(p.toString().indexOf(str)>=0){
					isAccepted = true;
				} else {
					if(p.toString().toUpperCase().indexOf(str.toUpperCase())>=0){
						isAccepted = true;
					}
				}
			}
		}
		
		return isAccepted;
	}
	
	public static int getStyleDefinitionIndex(Style style, StyleDefinitionsPart stylePart) {
		int index = 0;
		
		try {
			for(; index<stylePart.getContents().getStyle().size(); index++){
				if(stylePart.getContents().getStyle().get(index).getStyleId().equals(style.getStyleId())){
					break;
				} else {
				}
			}
			if(index>=stylePart.getContents().getStyle().size()){
				index=-1;
			}
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return index;
	}
	
	public static boolean isTitle(P p, StyleDefinitionsPart stylePart) {
		boolean isTitle = false;
		
		//ACM
		int fontSZ = 36; 
		String fontName = "Helvetica";
		String fontAlternative = "Times New Roman";
		boolean isBold = true; 
		boolean isItalic = false;
		
		if(sourceFormat==null
				&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)){
			isTitle = true;
			sourceFormat = SOURCEFORMAT_ACM;
		}
		
		//Elsevier
		if(!isTitle){
			fontSZ = 34;
			fontName = "Times New Roman";
			isBold = false;
			isItalic = false;
			if(sourceFormat==null
					&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)){
				isTitle = true;
				sourceFormat = SOURCEFORMAT_ELSEVIER;
			}
		}
		
		//IEEE
		if(!isTitle){
			fontSZ = 48;
			fontName = "Times New Roman";
			isBold = false;
			isItalic = false;
			if(sourceFormat==null
					&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)){
				isTitle = true;
				sourceFormat = SOURCEFORMAT_IEEE;
			}
		}
		//Springer
		if(!isTitle){
			fontSZ = 28;
			fontName = "Times";
			fontAlternative = "Times New Roman";
			isBold = true;
			isItalic = false;
			if(sourceFormat==null
					&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null||isAcceptedFontName(p, stylePart, fontAlternative, false))
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)){
				isTitle = true;
				sourceFormat = SOURCEFORMAT_SPRINGER;
			}
		}
		return isTitle;
	}
	
	public static boolean isTitleWithOld(P p, StyleDefinitionsPart stylePart) {
		boolean isTitle = false;
		if(p.getPPr()!=null){
			PPr pPr = p.getPPr();
			if(pPr.getPStyle()!=null 
					&& (pPr.getPStyle().getVal().equals("ArticleTitle")			// ACM
							|| pPr.getPStyle().getVal().equals("Titledocument")	// ACM
							|| pPr.getPStyle().getVal().equals("Title")			// IEEE
							|| pPr.getPStyle().getVal().equals("papertitle"))){	// Springer
				isTitle = true;
			} else if(pPr.getRPr()!=null && pPr.getRPr().getSz()!=null && pPr.getRPr().getSz().getVal().intValue() == 36	// ACM
					|| pPr.getRPr()!=null && pPr.getRPr().getSz()!=null && pPr.getRPr().getSz().getVal().intValue() == 48	// IEEE
					|| pPr.getRPr()!=null && pPr.getRPr().getSz()!=null && pPr.getRPr().getSz().getVal().intValue() == 28	// Springer
					){
				isTitle = true;
			} else if(pPr.getPStyle()!=null) {
				String styleId = pPr.getPStyle().getVal();
				Style style = stylePart.getStyleById(styleId);
				if(style!=null && style.getRPr()!=null && style.getRPr().getSz()!=null && style.getRPr().getSz().getVal()!=null){
					if(style.getRPr().getSz().getVal().intValue()== 36 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){		// ACM Title
						isTitle = true;
					} else if(style.getRPr().getSz().getVal().intValue()== 48 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){// IEEE Title
						isTitle = true;
					} else if(style.getRPr().getSz().getVal().intValue()== 28 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){// Springer Title
						isTitle = true;
					}
				}
			} else {
				
			}
		}
		return isTitle;
	}
	
	public static boolean isAffliation(P p, StyleDefinitionsPart stylePart){
		boolean isAffliation = false;
		
		return isAffliation;
	}
	
	public static boolean isAbstract(P p, StyleDefinitionsPart stylePart){
		boolean isAbstract = false;
		//ACM
		int fontSZ = 24; 
		String fontName = "Helvetica"; 
		boolean isBold = true; 
		boolean isItalic = false;
				
		if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& !isHasPNum(p, stylePart)
				&& isPTextStartWithString(p, "ABSTRACT", false)){
			isAbstract = true;
			sourceFormat = SOURCEFORMAT_ACM;
		}
		
		//Elsevier
		if(!isAbstract){
			fontSZ = 18;
			fontName = "Times New Roman";
			isBold = true;
			isItalic = false;
			String alignment = "left";
			if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& !isHasPNum(p, stylePart)
					&& isPTextStartWithString(p, "Abstract", false)){
				isAbstract = true;
				sourceFormat = SOURCEFORMAT_ELSEVIER;
			}
		}
				
		//IEEE
		if(!isAbstract){
			fontSZ = 18;
			fontName = "Times New Roman";
			isBold = true;
			isItalic = true;
			if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& !isHasPNum(p, stylePart)
					&& isPTextStartWithString(p, "Abstract", true)
					&& (isPTextIncludeString(p, ""+((char)45), true)
							||isPTextIncludeString(p, ""+((char)8212), true))){
				isAbstract = true;
				sourceFormat = SOURCEFORMAT_IEEE;
			}
		}

		//Springer
		if(!isAbstract){
			fontSZ = 18;
			fontName = "Times";
			isBold = true;
			isItalic = false;
			if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& !isHasPNum(p, stylePart)
					&& isPTextStartWithString(p, "Abstract", false)){
				isAbstract = true;
				sourceFormat = SOURCEFORMAT_SPRINGER;
			}
		}
		
		return isAbstract;
	}
	
	public static boolean isAuthors(P p, StyleDefinitionsPart stylePart){
		boolean isAuthors = false;
		
		if(OOXMLConvertTool.getPText(p).length()<=0 
				|| OOXMLConvertTool.getPText(p).replaceAll(" ", "").toLowerCase().indexOf("abstract")==0
				|| OOXMLConvertTool.getPText(p).replaceAll(" ", "").toLowerCase().indexOf("keyword")==0
				|| OOXMLConvertTool.getPText(p).replaceAll(" ", "").length()<=0){
			return isAuthors;
		}
		
		// ACM
		int fontSZ = 24; 
		String fontName = "Helvetica"; 
		boolean isBold = false; 
		boolean isItalic = false;
		String alignment = "center";
		if(SOURCEFORMAT_ACM.equals(sourceFormat)
				&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& !isHasPNum(p, stylePart)
				){
			isAuthors = true;
		}
		// IEEE
		fontSZ = 22;
		fontName = "Times New Roman";
		isBold = false;
		isItalic = false;
		alignment = "center";
		if(SOURCEFORMAT_IEEE.equals(sourceFormat)
				&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& !isHasPNum(p, stylePart)
				){
			isAuthors = true;
		}
		// Springer
		fontSZ = 20;
		fontName = "Times";
		isBold = false;
		isItalic = false;
		alignment = "center";
		if(SOURCEFORMAT_SPRINGER.equals(sourceFormat)
				&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& !isHasPNum(p, stylePart)
				){
			isAuthors = true;
		}
		// Elsevier
		fontSZ = 26;
		fontName = "Times New Roman";
		isBold = false;
		isItalic = false;
		alignment = "center";
		if(SOURCEFORMAT_ELSEVIER.equals(sourceFormat)
				&& OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& !isHasPNum(p, stylePart)
				){
			isAuthors = true;
		}
		return isAuthors;
	}
	
	public static boolean isAbstractOld(P p, StyleDefinitionsPart stylePart){
		boolean isAbstract = false;
		if(p.getPPr()!=null){
			PPr pPr = p.getPPr();
			if(pPr.getPStyle()!=null){
				if(pPr.getPStyle()!=null
						&& (pPr.getPStyle().getVal().equals("Abstract")			// ACM // IEEE
								|| pPr.getPStyle().getVal().equals("abstract"))	// Springer
						|| p.toString().toUpperCase().startsWith("ABSTRACT")){	
					isAbstract = true;
				} else if(pPr.getPStyle()!=null) {
					String styleId = pPr.getPStyle().getVal();
					Style style = stylePart.getStyleById(styleId);
					if(style!=null && style.getRPr()!=null && style.getRPr().getSz()!=null && style.getRPr().getSz().getVal()!=null){
						if(style.getRPr().getSz().getVal().intValue()== 24 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){		// ACM abstract
							isAbstract = true;
						} else if(style.getRPr().getSz().getVal().intValue()== 18 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){	// IEEE abstract
							isAbstract = true;
						} else if(style.getRPr().getSz().getVal().intValue()== 28 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){// Springer abstract
							// no keyword
						}
					}
					
				} else {
					
				}
			}
		}
		return isAbstract;
	}
	
	public static boolean isKeywords(P p, StyleDefinitionsPart stylePart) {
		boolean isKeywords = false;
		
		// ACM Keywords
		int fontSZ = 24; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		
		if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
				&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
				&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& !isHasPNum(p, stylePart)
				&& isPTextStartWithString(p, "Keyword", false)){
			isKeywords = true;
		}

		// Elsevier Keywords
		if(!isKeywords){
			fontSZ = 18; 
			fontName = "Times New Roman"; 
			isBold = false;
			isItalic = true;

			if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& !isHasPNum(p, stylePart)
					&& isPTextStartWithString(p, "Keyword", false)){
				isKeywords = true;
			}
		}
		
		// IEEE Keywords
		if(!isKeywords) {
			fontSZ = 18; 
			fontName = "Times New Roman"; 
			isBold = true;
			isItalic = true;

			if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& !isHasPNum(p, stylePart)
					&& (isPTextIncludeString(p, ""+((char)45), true)
							||isPTextIncludeString(p, ""+((char)8212), true))){
				isKeywords = true;
			}
		}
		
		// Springer Keywords
		if(!isKeywords) {
			fontSZ = 18; 
			fontName = "Times New Roman"; 
			isBold = true;
			isItalic = false;

			if(OOXMLConvertTool.isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& (OOXMLConvertTool.isAcceptedFontName(p, stylePart, fontName, false)||fontName==null)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& !isHasPNum(p, stylePart)
					&& isPTextStartWithString(p, "Keyword", false)){
				isKeywords = true;
			}
		}
		
		return isKeywords;
	}
	
	public static boolean isKeywordsOld(P p, StyleDefinitionsPart stylePart){
		boolean isKeywords = false;
		
		if(p.getPPr()!=null){
			PPr pPr = p.getPPr();
			if(p.getPPr().getPStyle()!=null && (p.getPPr().getPStyle().getVal().equals("keywords")	// Springer
						|| p.getPPr().getPStyle().getVal().equals("KeyWords")	// ACM	
						|| (p.getPPr().getPStyle().getVal().equals("abstract") && p.toString().toUpperCase().startsWith("KEYWORDS"))
						)){
					isKeywords = true;
			} else if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSz()!=null) {
				// ACM
				if(p.getPPr().getRPr().getSz().getVal().intValue()==24){
					if(p.getPPr().getRPr().getB()!=null && p.getPPr().getRPr().getB().isVal() && p.toString().equals("Keywords")){//ACM
						isKeywords = true;
					}
				} else if(p.getPPr().getRPr().getSz().getVal().intValue()==18) {
					if(p.getPPr().getRPr().getB()!=null && p.getPPr().getRPr().getB().isVal() && (p.toString().startsWith("Keywords—")||p.toString().toUpperCase().startsWith("KEYWORDS"))){//IEEE
						isKeywords = true;
					}
				}
				
			} else if(pPr.getPStyle()!=null) {
				String styleId = pPr.getPStyle().getVal();
				Style style = stylePart.getStyleById(styleId);
				if(style!=null && style.getRPr()!=null && style.getRPr().getSz()!=null && style.getRPr().getSz().getVal()!=null){
					if(style.getRPr().getSz().getVal().intValue()== 24 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){		// ACM keyword
						isKeywords = true;
					} else if(style.getRPr().getSz().getVal().intValue()== 18 && style.getRPr().getB()!=null && style.getRPr().getB().isVal() && style.getRPr().getI()!=null && style.getRPr().getI().isVal()){// IEEE keyword
						isKeywords = true;
					} else if(style.getRPr().getSz().getVal().intValue()== 28 && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){// Springer Title
						// no keyword
					}
				}
			} else {
				
			}
		}
		
		return isKeywords;
	}
	
	public static boolean isSectionHeader1(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart) {
		boolean isSectionHeader1 = false;
		
		// ACM Heading1
		int fontSZ = 24; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		boolean isCaps = true;
		boolean isSmallCaps = false;
		String alignment = "left";
		String pattern = "^\\d+.";
		String numberPattern = "%1.";
		NumberFormat numFm = NumberFormat.DECIMAL;
		
		if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& isAcceptedFontName(p, stylePart, fontName, false)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& isHasPNum(p, stylePart)
				&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
				&& !isEqualSpecialText(p)
				){
			isSectionHeader1 = true;
		}
		
		// Elsevier Heading1
		if(!isSectionHeader1){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = true; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			pattern = "^\\d+.";
			numberPattern = "%1.";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
					&& !isEqualSpecialText(p)
					){
				isSectionHeader1 = true;
			}
		}
		
		// IEEE Heading1
		if(!isSectionHeader1){
			fontSZ = 20;
			fontName = "Times New Roman";
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = true;
			alignment = "center";
			pattern = "^\\d+.";
			numberPattern = "%1.";
			numFm = NumberFormat.UPPER_ROMAN; 
			if((isAcceptedFontSZ(p, fontSZ, stylePart, false)||true)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
					&& !isEqualSpecialText(p)
					){
				isSectionHeader1 = true;
			}
		}
		
		// Springer Heading1
		if(!isSectionHeader1){
			fontSZ = 24;
			fontName = "Times"; 
			isBold = true; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			pattern = "^\\d+";
			numberPattern = "%1";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
					&& !isEqualSpecialText(p)
					){
				isSectionHeader1 = true;
			}
		}
		
		return isSectionHeader1;
	}
	
	public static boolean isSectionHeader2(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart) {
		boolean isSectionHeader2 = false;
		// ACM Heading2
		int fontSZ = 24; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		boolean isCaps = false;
		boolean isSmallCaps = false;
		String alignment = "left";
		String pattern = "^\\d+.\\d+";
		String numberPattern = "%1.%2";
		NumberFormat numFm = NumberFormat.DECIMAL;
		if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
				&& isAcceptedFontName(p, stylePart, fontName, true)
				&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& isHasPNum(p, stylePart)
				&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)){
			isSectionHeader2 = true;
		}
		
		// Elsevier Heading2
		if(!isSectionHeader2){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			pattern = "^\\d+.\\d+.";
			numberPattern = "%1.%2.";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)){
				isSectionHeader2 = true;
			}
		}
		
		// IEEE Heading2
		if(!isSectionHeader2){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			pattern = "^[A-Z].{1,1}?";
			numberPattern = "%2.";
			numFm = NumberFormat.UPPER_LETTER;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)){
				isSectionHeader2 = true;
			}
		}
		
		// Springer Heading2
		if(!isSectionHeader2){
			fontSZ = 20; 
			fontName = "Times"; 
			isBold = true; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^\\d+.\\d+";
			numberPattern = "%1.%2";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)){
				isSectionHeader2 = true;
			}
		}
		
		return isSectionHeader2;
	}
	
	public static boolean isSectionHeader3(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart) {
		boolean isSectionHeader3 = false;

		// ACM Heading3
		int fontSZ = 22; 
		String fontName = "Times New Roman";
		boolean isBold = false; 
		boolean isItalic = true;
		boolean isCaps = false;
		boolean isSmallCaps = false;
		String alignment = "left";
		String pattern = "^\\d+.\\d+.\\d+";
		String numberPattern = "%1.%2.%3";
		NumberFormat numFm = NumberFormat.DECIMAL;
		if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
				&& isAcceptedFontName(p, stylePart, fontName, true)
				&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
				&& (isAcceptedFontAlignment(p, stylePart, alignment)||isAcceptedFontAlignment(p, stylePart, "both"))
				&& isHasPNum(p, stylePart)
				&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)){
			isSectionHeader3 = true;
		}
		
		// Elsevier Heading3
		if(!isSectionHeader3){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
//			alignment = "left";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
//					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& !isHasPNum(p, stylePart)
					&& !isStyleTypeParagraph(p, stylePart)
					&& !isAllRInSameStyle(p, stylePart)
					){
				isSectionHeader3 = true;
			}
		}
		
		// IEEE heading3
		if(!isSectionHeader3){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^\\d+\\)";
			numberPattern = "%3)";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
//					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isHasPNum(p, stylePart)
					&& !isStyleTypeParagraph(p, stylePart)
					&& !isAllRInSameStyle(p, stylePart)
					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
					){
				isSectionHeader3 = true;
			}
		}
		
		// Springer Heading3
		if(!isSectionHeader3){
			fontSZ = 20; 
			fontName = "Times"; 
			isBold = true; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
//					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& !isHasPNum(p, stylePart)
					&& !isStyleTypeParagraph(p, stylePart)
					&& !isAllRInSameStyle(p, stylePart)
					){
				isSectionHeader3 = true;
			}
		}
		return isSectionHeader3;
	}
	
	public static boolean isSectionHeader4(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart) {
		boolean isSectionHeader4 = false;

		// ACM Heading4
		int fontSZ = 22; 
		String fontName = "Times New Roman"; 
		boolean isBold = false; 
		boolean isItalic = true;
		boolean isCaps = false;
		boolean isSmallCaps = false;
		String alignment = "left";
		String pattern = "^\\d+.\\d+.\\d+.\\d+";
		String numberPattern = "%1.%2.%3.%4";
		NumberFormat numFm = NumberFormat.DECIMAL;
		if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& isAcceptedFontName(p, stylePart, fontName, false)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
				&& (isAcceptedFontAlignment(p, stylePart, alignment) || isAcceptedFontAlignment(p, stylePart, "both"))
				&& isHasPNum(p, stylePart)
				&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)){
			isSectionHeader4 = true;
		}
		
		// Elsevier Heading4
		if(!isSectionHeader4){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
//			alignment = "left";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& !isHasPNum(p, stylePart)
					&& !isStyleTypeParagraph(p, stylePart)
					&& !isAllRInSameStyle(p, stylePart)
					){
				isSectionHeader4 = true;
			}
		}
		
		// IEEE Heading4
		if(!isSectionHeader4){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^[a-z]+\\)";
			numberPattern = "%4)";
			numFm = NumberFormat.LOWER_LETTER;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
					&& (isAcceptedFontAlignment(p, stylePart, alignment) || isAcceptedFontAlignment(p, stylePart, "left"))
					&& isHasPNum(p, stylePart)
//					&& !isStyleTypeParagraph(p, stylePart)
					&& !isAllRInSameStyle(p, stylePart)
//					&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
					){
				isSectionHeader4 = true;
			}
		}
		
		// Springer Heading4
		if(!isSectionHeader4){
			fontSZ = 20; 
			fontName = "Times"; 
			isBold = false; 
			isItalic = true;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, true)
					&& (isBold?isAcceptedFontBold(p, stylePart, true):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, true):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
					&& (isAcceptedFontAlignment(p, stylePart, alignment) || isAcceptedFontAlignment(p, stylePart, "both"))
					&& !isHasPNum(p, stylePart)
					&& !isStyleTypeParagraph(p, stylePart)
					&& !isAllRInSameStyle(p, stylePart)
					){
				isSectionHeader4 = true;
			}
		}
		return isSectionHeader4;
	}
	/*
	 * check if the font size is suitable
	 */
	public static boolean isAcceptedFontSZ(P p, int fontSZ, StyleDefinitionsPart stylePart, boolean isDeepChk) {
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null) {
			if(p.getContent()!=null && p.getContent().size()>0 && isDeepChk) {
				Style chkStyle = new Style();
				chkStyle.setRPr(new RPr());
				chkStyle.getRPr().setSz(new HpsMeasure());
				chkStyle.getRPr().getSz().setVal(new BigInteger(fontSZ+""));
				isAccepted = chkStyleInR(p, chkStyle, stylePart);
			}
			if(isAccepted){
				return true;
			}
			if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSz()!=null && p.getPPr().getRPr().getSz().getVal()!=null){
				if(p.getPPr().getRPr().getSz().getVal().intValue()==fontSZ){
					isAccepted = true;
				}
			} else if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSzCs()!=null && p.getPPr().getRPr().getSzCs().getVal()!=null){
				if(p.getPPr().getRPr().getSzCs().getVal().intValue()==fontSZ){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				boolean isStyleHas= false, isBaseOnStyleHas = false, isLinkedStyleHas = false;
				if(style!=null){
					if(style.getRPr()!=null && style.getRPr().getSz()!=null && style.getRPr().getSz().getVal()!=null){
						if(style.getRPr().getSz().getVal().intValue()== fontSZ){
							isAccepted = true;
						}
					} else {
						isStyleHas = true;
					}
					if(!isAccepted){
						if(style.getBasedOn()!=null){
							Style basedOnStyle = stylePart.getStyleById(style.getBasedOn().getVal());
							if(basedOnStyle.getRPr()!=null && basedOnStyle.getRPr().getSz()!=null && basedOnStyle.getRPr().getSz().getVal()!=null){
								if(basedOnStyle.getRPr().getSz().getVal().intValue()== fontSZ){
									isAccepted = true;
								}
							} else {
								isBaseOnStyleHas = true;
							}
							if(!isAccepted){
								if(basedOnStyle.getLink()!=null){
									Style linkedStyle = stylePart.getStyleById(basedOnStyle.getLink().getVal());
									if(linkedStyle.getRPr()!=null && linkedStyle.getRPr().getSz()!=null && linkedStyle.getRPr().getSz().getVal()!=null) {
										if(linkedStyle.getRPr().getSz().getVal().intValue() == fontSZ){
											isAccepted = true;
										}
									} else {
										isLinkedStyleHas = true;
									}
								}
							}
						}
					}
					if(!isAccepted){
						if(style.getLink()!=null){
							Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
							if(linkedStyle.getRPr()!=null && linkedStyle.getRPr().getSz()!=null && linkedStyle.getRPr().getSz().getVal()!=null) {
								if(linkedStyle.getRPr().getSz().getVal().intValue() == fontSZ){
									isAccepted = true;
								}
							} else {
								isLinkedStyleHas = true;
							}
						}
					}
//					if(isStyleHas && isBaseOnStyleHas && isLinkedStyleHas && !AppEnvironment.getInstance().isStrict()){
//						isAccepted = true;
//					}
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		if(!isAccepted){
			if(stylePart.getDefaultParagraphStyle()!=null 
					&& stylePart.getDefaultParagraphStyle().getPPr()!=null
					&& stylePart.getDefaultParagraphStyle().getPPr().getRPr()!=null){
				if(stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz().getVal()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz().getVal().intValue()==fontSZ
						){
					isAccepted = true;
				}
			}
		}
		return isAccepted;
	}
	
	/*
	 * check if the font name is suitable
	 */
	private static boolean isAcceptedFontName(P p, StyleDefinitionsPart stylePart, String fontName, boolean isDeepChk) {
		boolean isAccepted = false;
		if(!AppEnvironment.getInstance().isActiveTestMode()){
			return true;
		}
		if(p!=null && p.getPPr()!=null){
			if(p.getContent()!=null && p.getContent().size()>0 && isDeepChk){
				for(Object o : p.getContent()){
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null){
							if(r.getRPr().getRFonts()!=null 
									&& (r.getRPr().getRFonts().getAscii()!=null 
										&& r.getRPr().getRFonts().getAscii().equals(fontName)
									|| r.getRPr().getRFonts().getCs()!=null 
										&& r.getRPr().getRFonts().getCs().equals(fontName))
									){
								isAccepted = true;
								break;
							} else if(r.getRPr().getRStyle()!=null){
								Style style = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
								if(style != null && style.getName()!=null && style.getName().equals(fontName)){
									isAccepted = true;
									break;
								}
							}
						}
					}
				}
			} else if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getRFonts()!=null
					&& p.getPPr().getRPr().getRFonts().getAscii()!=null) {
				
				if(p.getPPr().getRPr().getRFonts().getAscii()!=null){
					if(p.getPPr().getRPr().getRFonts().getAscii().equals(fontName)){
						isAccepted = true;
					}
				} else if(p.getPPr().getRPr().getRFonts().getCs()!=null){
					if(p.getPPr().getRPr().getRFonts().getCs().equals(fontName)){
						isAccepted = true;
					}
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getRPr()!=null && style.getRPr().getRFonts()!=null){
						RFonts rFonts = style.getRPr().getRFonts();
						if(rFonts.getAscii()!=null && rFonts.getAscii().equals(fontName)
								|| rFonts.getCs()!=null && rFonts.getCs().equals(fontName)){
							isAccepted = true;
						}
					} 
					
					if(!isAccepted){
						if(style.getBasedOn()!=null){
							Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
							if(styleBasedOn!=null && styleBasedOn.getRPr()!=null && styleBasedOn.getRPr().getRFonts()!=null){
								if(styleBasedOn.getRPr().getRFonts().getAscii().equals(fontName)){
									isAccepted = true;
								}
							}
						}
					}
					
					if(!isAccepted){
						if(style.getLink()!=null){
							Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
							if(linkedStyle!=null && linkedStyle.getRPr()!=null && linkedStyle.getRPr().getRFonts()!=null){
								if(linkedStyle.getRPr().getRFonts().getAscii().equals(fontName)){
									isAccepted = true;
								}
							}
						}
					}
					
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		if(!isAccepted){
			if(stylePart.getDefaultParagraphStyle()!=null 
					&& stylePart.getDefaultParagraphStyle().getPPr()!=null
					&& stylePart.getDefaultParagraphStyle().getPPr().getRPr()!=null){
				if(stylePart.getDefaultParagraphStyle().getPPr().getRPr().getRFonts()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getRFonts()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getRFonts().getAscii()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getRFonts().getAscii().equals(fontName)
						){
					isAccepted = true;
				}
			}
		}
		return isAccepted;
	}
	
	/*
	 * check if font bold
	 */
	public static boolean isAcceptedFontBold(P p, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null){
			if(isDeepChk){
				Style style = new Style();
				style.setRPr(new RPr());
				style.getRPr().setB(new BooleanDefaultTrue());
				isAccepted = isHasStyleInDeep(p, style, stylePart);
			}
			if(isAccepted){
				return true;
			}
			if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getB()!=null) {
				if(p.getPPr().getRPr().getB().isVal()){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getRPr()!=null && (style.getRPr().getB()!=null || style.getRPr().getBCs()!=null)){
						if(style.getRPr().getB()!=null && style.getRPr().getB().isVal() || style.getRPr().getBCs()!=null && style.getRPr().getBCs().isVal()){
							isAccepted = true;
						}
					}
					if(!isAccepted && style.getBasedOn()!=null){
						Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
						if(styleBasedOn!=null && styleBasedOn.getRPr()!=null 
								&& (styleBasedOn.getRPr().getB()!=null
								&& styleBasedOn.getRPr().getB().isVal()
								|| styleBasedOn.getRPr().getBCs()!=null
								&& styleBasedOn.getRPr().getBCs().isVal())) {
							isAccepted = true;
						}
						if(!isAccepted){
							if(styleBasedOn.getLink()!=null){
								Style linkedStyle = stylePart.getStyleById(styleBasedOn.getLink().getVal());
								if(linkedStyle.getRPr()!=null && (linkedStyle.getRPr().getB()!=null && linkedStyle.getRPr().getB().isVal() || linkedStyle.getRPr().getBCs()!=null && linkedStyle.getRPr().getBCs().isVal())) {
									isAccepted = true;
								}
							}
						}
					}
					if(!isAccepted && style.getLink()!=null){
						Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
						if(linkedStyle!=null && linkedStyle.getRPr()!=null 
								&& linkedStyle.getRPr().getB()!=null 
								&& linkedStyle.getRPr().getB().isVal()){
							isAccepted = true;
						}
					}
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		
		return isAccepted;
	}
	/*
	 * check if font italic
	 */
	public static boolean isAcceptedFontItalic(P p, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null){
			if(isDeepChk) {
				Style style = new Style();
				style.setRPr(new RPr());
				style.getRPr().setI(new BooleanDefaultTrue());
				isAccepted = isHasStyleInDeep(p, style, stylePart);
			}
			if(isAccepted){
				return true;
			}
			if(p.getPPr().getRPr()!=null && (p.getPPr().getRPr().getI()!=null||p.getPPr().getRPr().getICs()!=null)) {
				if(p.getPPr().getRPr().getI()!=null && p.getPPr().getRPr().getI().isVal()
						|| p.getPPr().getRPr().getICs()!=null && p.getPPr().getRPr().getICs().isVal() ){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getRPr()!=null && style.getRPr().getI()!=null){
						if(style.getRPr().getI().isVal()){
							isAccepted = true;
						}
					}
					if(!isAccepted && style.getBasedOn()!=null){
						Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
						if(styleBasedOn!=null && styleBasedOn.getRPr()!=null 
								&& (styleBasedOn.getRPr().getI()!=null
								&& styleBasedOn.getRPr().getI().isVal()
								|| styleBasedOn.getRPr().getICs()!=null
								&& styleBasedOn.getRPr().getICs().isVal())) {
							isAccepted = true;
						}
						if(!isAccepted){
							if(styleBasedOn.getLink()!=null){
								Style linkedStyle = stylePart.getStyleById(styleBasedOn.getLink().getVal());
								if(linkedStyle.getRPr()!=null && (linkedStyle.getRPr().getI()!=null && linkedStyle.getRPr().getI().isVal())||linkedStyle.getRPr().getICs()!=null && linkedStyle.getRPr().getICs().isVal()) {
									isAccepted = true;
								}
							}
						}
					}
					if(!isAccepted && style.getLink()!=null){
						Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
						if(linkedStyle!=null && linkedStyle.getRPr()!=null 
								&& linkedStyle.getRPr().getI()!=null 
								&& linkedStyle.getRPr().getI().isVal()){
							isAccepted = true;
						}
					}
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		
		return isAccepted;
	}
	
	/*
	 * check if font with underline
	 */
	private static boolean isAcceptedFontWithUnderline(P p, StyleDefinitionsPart stylePart, UnderlineEnumeration underlineType, boolean isDeepChk){
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null){
			if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getU()!=null) {
				if(p.getPPr().getRPr().getU().getVal()!=null && p.getPPr().getRPr().getU().getVal().equals(underlineType)){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getRPr()!=null && style.getRPr().getU()!=null){
						if(style.getRPr().getU().getVal()!=null && style.getRPr().getU().getVal().equals(underlineType)){
							isAccepted = true;
						}
						if(!isAccepted && style.getBasedOn()!=null){
							Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
							if(styleBasedOn!=null && styleBasedOn.getRPr()!=null 
									&& styleBasedOn.getRPr().getU()!=null
									&& styleBasedOn.getRPr().getU().getVal()!=null
									&& styleBasedOn.getRPr().getU().getVal().equals(underlineType)) {
								isAccepted = true;
							}
						}
						if(!isAccepted && style.getLink()!=null){
							Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
							if(linkedStyle!=null && linkedStyle.getRPr()!=null 
									&& linkedStyle.getRPr().getU()!=null 
									&& linkedStyle.getRPr().getU().getVal()!=null
									&& linkedStyle.getRPr().getU().getVal().equals(linkedStyle)){
								isAccepted = true;
							}
						}
					} 
				}
			} else if(isDeepChk) {
				Style style = new Style();
				style.setRPr(new RPr());
				style.getRPr().setU(new U());
				style.getRPr().getU().setVal(underlineType);
				isAccepted = chkStyleInR(p, style, stylePart);
			}
		}
		
		return isAccepted;
	}
	/*
	 * check if font small Caps
	 */
	private static boolean isAcceptedFontSmallCaps(P p, StyleDefinitionsPart stylePart, boolean isDeepChk) {
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null){
			if(isDeepChk) {
				Style style = new Style();
				style.setRPr(new RPr());
				style.getRPr().setSmallCaps(new BooleanDefaultTrue());
				isAccepted = isHasStyleInDeep(p, style, stylePart);
			}
			if(isAccepted){
				return true;
			}
			if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSmallCaps()!=null) {
				if(p.getPPr().getRPr().getSmallCaps().isVal()){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getRPr()!=null && style.getRPr().getSmallCaps()!=null){
						if(style.getRPr().getSmallCaps().isVal()){
							isAccepted = true;
						}
						if(!isAccepted && style.getBasedOn()!=null){
							Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
							if(styleBasedOn!=null && styleBasedOn.getRPr()!=null 
									&& styleBasedOn.getRPr().getSmallCaps()!=null
									&& styleBasedOn.getRPr().getSmallCaps().isVal()) {
								isAccepted = true;
							}
						}
						if(!isAccepted && style.getLink()!=null){
							Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
							if(linkedStyle!=null && linkedStyle.getRPr()!=null 
									&& linkedStyle.getRPr().getSmallCaps()!=null 
									&& linkedStyle.getRPr().getSmallCaps().isVal()){
								isAccepted = true;
							}
						}
					} 
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		if(!isAccepted){
			if(stylePart.getDefaultParagraphStyle()!=null 
					&& stylePart.getDefaultParagraphStyle().getPPr()!=null
					&& stylePart.getDefaultParagraphStyle().getPPr().getRPr()!=null){
				if(stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSmallCaps()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSmallCaps()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSmallCaps().isVal()){
					isAccepted = true;
				}
			}
		}
		return isAccepted;
	}
	/*
	 * check if font is caps
	 */
	private static boolean isAcceptedFontCaps(P p, StyleDefinitionsPart stylePart, boolean isDeepChk) {
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null){
			if(p.toString().toUpperCase().equals(p.toString())){
				isAccepted = true;
			} else if(isDeepChk) {
				Style style = new Style();
				style.setRPr(new RPr());
				style.getRPr().setCaps(new BooleanDefaultTrue());
				isAccepted = isHasStyleInDeep(p, style, stylePart);
			}
			if(isAccepted){
				return true;
			}
			if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getCaps()!=null) {
				if(p.getPPr().getRPr().getCaps().isVal()){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getRPr()!=null && style.getRPr().getCaps()!=null){
						if(style.getRPr().getCaps().isVal()){
							isAccepted = true;
						}
					} 
					if(!isAccepted && style.getBasedOn()!=null){
						Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
						if(styleBasedOn!=null && styleBasedOn.getRPr()!=null 
								&& styleBasedOn.getRPr().getCaps()!=null
								&& styleBasedOn.getRPr().getCaps().isVal()) {
							isAccepted = true;
						}
					}
					if(!isAccepted && style.getLink()!=null){
						Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
						if(linkedStyle!=null && linkedStyle.getRPr()!=null 
								&& linkedStyle.getRPr().getCaps()!=null 
								&& linkedStyle.getRPr().getCaps().isVal()){
							isAccepted = true;
						}
					}
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		if(!isAccepted){
			if(stylePart.getDefaultParagraphStyle()!=null 
					&& stylePart.getDefaultParagraphStyle().getPPr()!=null
					&& stylePart.getDefaultParagraphStyle().getPPr().getRPr()!=null){
				if(stylePart.getDefaultParagraphStyle().getPPr().getRPr().getCaps()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getCaps()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getCaps().isVal()){
					isAccepted = true;
				}
			}
		}
		return isAccepted;
	}
	/*
	 * check font alignment
	 */
	private static boolean isAcceptedFontAlignment(P p, StyleDefinitionsPart stylePart, String alignment) {
		boolean isAccepted = false;
		
		boolean isNotDefined = true;
		
		if(p!=null && p.getPPr()!=null){
			if(p.getPPr().getJc()!=null && p.getPPr().getJc().getVal()!=null){
				if(p.getPPr().getJc().getVal().value().equals(alignment)){
					isAccepted = true;
				}
			} else if(p.getPPr().getTextAlignment()!=null) {
				TextAlignment textAlignment = p.getPPr().getTextAlignment();
				if(textAlignment.getVal().equals(alignment)){
					isAccepted = true;
				}
			} else if(p.getPPr().getPStyle()!=null){
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null){
					if(style.getPPr()!=null && style.getPPr().getJc()!=null && style.getPPr().getJc().getVal()!=null){
						if(style.getPPr().getJc().getVal().value().equals(alignment)){
							isAccepted = true;
						} else {
							isNotDefined = false;
						}
					} else if(style.getPPr()!=null && style.getPPr().getTextAlignment()!=null){
						if(style.getPPr().getTextAlignment().getVal().equals(alignment)){
							isAccepted = true;
						} else {
							isNotDefined = false;
						}
					} else if(!isAccepted && style.getBasedOn()!=null){
						Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
						if(styleBasedOn!=null){
							if(styleBasedOn.getPPr()!=null
									&& styleBasedOn.getPPr().getJc()!=null
									&& styleBasedOn.getPPr().getJc().getVal()!=null){
								if(styleBasedOn.getPPr().getJc().getVal().value().equals(alignment)){
									isAccepted = true;
								} else {
									isNotDefined = false;
								}
							} else if(styleBasedOn.getPPr()!=null 
									&& styleBasedOn.getPPr().getTextAlignment()!=null){
								if(styleBasedOn.getPPr().getTextAlignment().getVal().equals(alignment)){
									isAccepted = true;
								} else {
									isNotDefined = false;
								}
							} else if(!isAccepted && style.getLink()!=null){
								Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
								if(linkedStyle!=null){
									if(linkedStyle.getPPr()!=null
											&& linkedStyle.getPPr().getJc()!=null
											&& linkedStyle.getPPr().getJc().getVal()!=null){
										if(linkedStyle.getPPr().getJc().getVal().value().equals(alignment)){
											isAccepted = true;
										} else {
											isNotDefined = false;
										}
									}
									if(linkedStyle.getPPr()!=null 
											&& linkedStyle.getPPr().getTextAlignment()!=null){
										if(linkedStyle.getPPr().getTextAlignment().getVal().equals(alignment)){
											isAccepted = true;
										}
									}
								}
							}
						}
					} 
				}
			} else {
//				if(!AppEnvironment.getInstance().isStrict()){
//					isAccepted = true;
//				}
			}
		}
		if(!isAccepted){
			if(stylePart.getDefaultParagraphStyle()!=null 
					&& stylePart.getDefaultParagraphStyle().getPPr()!=null
					&& stylePart.getDefaultParagraphStyle().getPPr().getJc()!=null){
				if(stylePart.getDefaultParagraphStyle().getPPr().getJc().getVal()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getJc().getVal()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getJc().getVal().value().equals(alignment)){
					isAccepted = true;
				}
			} else {
				if("left".equals(alignment) && isNotDefined){
					isAccepted = true;
				}
			}
		}
		return isAccepted;
	}
	
	private static boolean isB(RPr rPr, Style style){
		boolean is1 = false;
		boolean is2 = false;
		
		if(rPr!=null && rPr.getB()!=null && rPr.getB().isVal()){
			is1 = true;
		}
		
		if(style!=null && style.getRPr()!=null && style.getRPr().getB()!=null && style.getRPr().getB().isVal()){
			is2 = true;
		}
		
		if(!is1 && !is2){
			if(rPr==null || (rPr!=null && rPr.getB()==null) || style==null || (style!=null && style.getRPr()==null) || (style!=null && style.getRPr()!=null && style.getRPr().getB()==null)){
				return false;
			}
		}
		
		return is1!=is2;
	}
	
	private static boolean isI(RPr rPr, Style style){
		boolean is1 = false;
		boolean is2 = false;
		
		if(rPr!=null && rPr.getI()!=null && rPr.getI().isVal()){
			is1 = true;
		}
		
		if(style!=null && style.getRPr()!=null && style.getRPr().getI()!=null && style.getRPr().getI().isVal()){
			is2 = true;
		}
		
		if(!is1 && !is2){
			if(rPr==null || (rPr!=null && rPr.getI()==null) || style==null || (style!=null && style.getRPr()==null) || (style!=null && style.getRPr()!=null && style.getRPr().getI()==null)){
				return false;
			}
		}
		
		return is1!=is2;
	}
	
	private static boolean isSmallCaps(RPr rPr, Style style){
		boolean is1 = false;
		boolean is2 = false;
		
		if(rPr!=null && rPr.getSmallCaps()!=null && rPr.getSmallCaps().isVal()){
			is1 = true;
		}
		
		if(style!=null && style.getRPr()!=null && style.getRPr().getSmallCaps()!=null && style.getRPr().getSmallCaps().isVal()){
			is2 = true;
		}
		if(!is1 && !is2){
			if(rPr==null || (rPr!=null && rPr.getSmallCaps()==null) || style==null || (style!=null && style.getRPr()==null) || (style!=null && style.getRPr()!=null && style.getRPr().getSmallCaps()==null)){
				return false;
			}
		}
		return is1!=is2;
	}
	
	private static boolean isCaps(RPr rPr, Style style){
		boolean is1 = false;
		boolean is2 = false;
		
		if(rPr!=null && rPr.getCaps()!=null){
			if(rPr.getCaps().isVal()){
				is1 = true;
			}
		}
		
		if(style!=null && style.getRPr()!=null && style.getRPr().getCaps()!=null){
			if(style.getRPr().getCaps().isVal()){
				is2 = true;
			}
		}
		
		if(!is1 && !is2){
			if(rPr==null || (rPr!=null && rPr.getCaps()==null) || style==null || (style!=null && style.getRPr()==null) || (style!=null && style.getRPr()!=null && style.getRPr().getCaps()==null)){
				return false;
			}
		}
		
		return is1!=is2;
	}
	
	public static boolean isAllRInSameStyle(P p, StyleDefinitionsPart stylePart){
		boolean isAccepted = true;
		
		List<Object> pContents = p.getContent();
		RPr firstRPr = null;
		Style firstRStyle = null;
		boolean isFirstB = false, isFirstI = false, isFirstCaps = false, isFirstSmallCaps = false;
		int index = -1;
		for (int i = 0; i < pContents.size(); i++) {
			Object o = pContents.get(i);
			if(o instanceof R){
				R r = (R)o;
				String rText = OOXMLConvertTool.getRText(r);
				if(rText!=null && rText.length()>0){
					
				} else {
					continue;
				}
				if(index<0){
					firstRPr = r.getRPr();
					firstRStyle = (firstRPr!=null && firstRPr.getRStyle()!=null)?stylePart.getStyleById(firstRPr.getRStyle().getVal()):null;
					isFirstB = isB(firstRPr, firstRStyle);
					isFirstI = isI(firstRPr, firstRStyle);
					isFirstSmallCaps = isSmallCaps(firstRPr, firstRStyle);
					isFirstCaps = isCaps(firstRPr, firstRStyle);
					index = i;
				} else {
					Style rStyle = (r.getRPr()!=null && r.getRPr().getRStyle()!=null)?stylePart.getStyleById(r.getRPr().getRStyle().getVal()):null;
					boolean isB = isB(r.getRPr(), rStyle);
					boolean isI = isI(r.getRPr(), rStyle);
					boolean isSmallCaps = isSmallCaps(r.getRPr(), rStyle);
					boolean isCaps = isCaps(r.getRPr(), rStyle);
					if(isB && isFirstB || !isB && !isFirstB){
						
					} else {
						return false;
					}
					if(isI && isFirstI || !isI && !isFirstI){
						
					} else {
						return false;
					}
					if(isSmallCaps && isFirstSmallCaps || !isSmallCaps && !isFirstSmallCaps){
						
					} else {
						return false;
					}
					if(isCaps && isFirstCaps || !isCaps && !isFirstCaps){
						
					} else {
						return false;
					}
				}
				
			}
		}
		
		return isAccepted;
	}
	/*
	public static boolean isAllRInSameStyle(P p, StyleDefinitionsPart stylePart){
		boolean isAccepted = true;
		
		List<Object> pContents = p.getContent();
		RPr rPr = null;
		int index = -1;
		for (int i = 0; i < pContents.size(); i++) {
			Object o = pContents.get(i);
			if(o instanceof R){
				R r = (R)o;
				String rText = OOXMLConvertTool.getRText(r);
				if(rText!=null && rText.length()>0){
					
				} else {
					continue;
				}
				if(rPr==null && index<0){
					rPr = r.getRPr();
					index = i;
				} else {
					Style rStyle = (r.getRPr()!=null && r.getRPr().getRStyle()!=null)?stylePart.getStyleById(r.getRPr().getRStyle().getVal()):null;
					Style firstRStyle = (rPr!=null && rPr.getRStyle()!=null)?stylePart.getStyleById(rPr.getRStyle().getVal()):null;
					
					if(r.getRPr()!=null && rPr!=null){
						if(rPr.getB()!=null && r.getRPr().getB()!=null){
							if(rPr.getB().isVal() && r.getRPr().getB().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getB()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getB()!=null 
										&& firstRStyle.getRPr().getB().isVal()){
									chk1 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(r.getRPr().getB()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getB()!=null
										&& rStyle.getRPr().getB().isVal()){
									chk2 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(rPr.getI()!=null && r.getRPr().getI()!=null){
							if(rPr.getI().isVal() && r.getRPr().getI().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getI()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getI()!=null 
										&& firstRStyle.getRPr().getI().isVal()){
									chk1 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(r.getRPr()==null || r.getRPr()!=null && r.getRPr().getI()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getI()!=null
										&& rStyle.getRPr().getI().isVal()){
									chk2 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(rPr.getCaps()!=null && r.getRPr().getCaps()!=null){
							if(rPr.getCaps().isVal() && r.getRPr().getCaps().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getCaps()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getCaps()!=null 
										&& firstRStyle.getRPr().getCaps().isVal()){
									chk1 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(r.getRPr().getCaps()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getCaps()!=null
										&& rStyle.getRPr().getCaps().isVal()){
									chk2 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(rPr.getSmallCaps()!=null && r.getRPr().getSmallCaps()!=null){
							if(rPr.getSmallCaps().isVal() && r.getRPr().getSmallCaps().isVal()){
								
							} else {
								isAccepted = false;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getSmallCaps()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getSmallCaps()!=null 
										&& firstRStyle.getRPr().getSmallCaps().isVal()){
									chk1 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(r.getRPr().getSmallCaps()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getSmallCaps()!=null
										&& rStyle.getRPr().getSmallCaps().isVal()){
									chk2 = true;
								} else {
									isAccepted = false;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								isAccepted = false;
								break;
							}
						}
						if(rPr.getRFonts()!=null && r.getRPr().getRFonts()!=null){
							if(rPr.getRFonts().getAscii()!=null && r.getRPr().getRFonts().getAscii()!=null
									&& rPr.getRFonts().getAscii().equals(r.getRPr().getRFonts().getAscii())){
								
							} else {
								isAccepted = false;
								break;
							}
						} else {
//							boolean chk1 = false, chk2 = false;
//							if(rPr.getRFonts()!=null && rPr.getRFonts().getAscii()==null ){
//								
//							}
						}
						if(rPr.getSz()!=null && r.getRPr().getSz()!=null){
							if(rPr.getSz().getVal()!=null && r.getRPr().getSz().getVal()!=null
									&& rPr.getSz().getVal().intValue()==r.getRPr().getSz().getVal().intValue()){
							} else if(rPr.getSz().getVal()==null || r.getRPr().getSz().getVal()==null) {
								
							} else {
								isAccepted = false;
								break;
							}
						} else {
							BigInteger chk1 = null;
							BigInteger chk2 = null;
							if(rPr.getSz()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null
										&& firstRStyle.getRPr().getSz()!=null){
									chk1 = firstRStyle.getRPr().getSz().getVal();
								} else {
									isAccepted = false;
									break;
								}
							}
							if(r.getRPr()!=null && r.getRPr().getSz()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getSz()!=null){
									chk2 = rStyle.getRPr().getSz().getVal();
								} else {
									isAccepted = false;
									break;
								}
							}
							if(chk1!=null && chk2!=null && chk1.intValue() == chk2.intValue()){
								
							} else {
								isAccepted = false;
								break;
							}
						}
					} else if(rPr==null && r.getRPr()==null){
						
					} else {
						isAccepted = false;
						break;
					}
				}
			}
		}
		
		return isAccepted;
	}*/
	
	public static boolean isAllRInSameStyle(P p, boolean test){
		boolean isAccepted = true;
		
		List<Object> pContents = p.getContent();
		Style style = new Style();
		style.setRPr(new RPr());
		boolean isDone = false;
		boolean isNull = false;
		for(Object o : pContents){
			if(o instanceof R){
				R r = (R)o;
				if(!isNull){
					if(r.getRPr()==null){
						isNull = true;
					}
				} else if(r.getRPr()!=null){
					if(r.getRPr().getB()==null && r.getRPr().getBCs()==null
							&& r.getRPr().getI()==null && r.getRPr().getICs()==null
							&& r.getRPr().getCaps()==null && r.getRPr().getSmallCaps()==null
							&& r.getRPr().getSz()==null && r.getRPr().getSzCs()==null){
						
					} else {
						isAccepted = false;
						break;
					}
				}
				if(r.getRPr()!=null && ! isDone){
					isDone = true;
					if(r.getRPr().getB()!=null && r.getRPr().getB().isVal()){
						style.getRPr().setB(new BooleanDefaultTrue());
					}
					if(r.getRPr().getI()!=null && r.getRPr().getI().isVal()){
						style.getRPr().setI(new BooleanDefaultTrue());
					}
					if(r.getRPr().getCaps()!=null && r.getRPr().getCaps().isVal()){
						style.getRPr().setCaps(new BooleanDefaultTrue());
					}
					if(r.getRPr().getSmallCaps()!=null && r.getRPr().getSmallCaps().isVal()){
						style.getRPr().setSmallCaps(new BooleanDefaultTrue());
					}
					if(r.getRPr().getRFonts()!=null){
						style.getRPr().setRFonts(r.getRPr().getRFonts());
					}
					if(r.getRPr().getSz()!=null){
						style.getRPr().setSz(r.getRPr().getSz());
					}
				} else if(r.getRPr()!=null) {
					if(r.getRPr().getB()!=null && style.getRPr().getB()!=null){
						if(r.getRPr().getB().isVal() 
								&& style.getRPr().getB().isVal()){
							
						} else {
							isAccepted = false;
							break;
						}
					} else {
						// ignore
					}
					if(r.getRPr().getI()!=null && style.getRPr().getI()!=null){
						if(r.getRPr().getI().isVal() && style.getRPr().getI().isVal()){
							
						} else {
							isAccepted = false;
							break;
						}
					} else {
						// ignore
					}
					if(r.getRPr().getCaps()!=null && style.getRPr().getCaps()!=null){
						if(r.getRPr().getCaps().isVal() && style.getRPr().getCaps().isVal()){
							
						} else {
							isAccepted = false;
							break;
						}
					} else {
						// ignore
					}
					if(r.getRPr().getSmallCaps()!=null && style.getRPr().getSmallCaps()!=null ){
						if(r.getRPr().getSmallCaps().isVal() && style.getRPr().getSmallCaps().isVal()){
							
						} else {
							isAccepted = false;
							break;
						}
					} else {
						// ignore
					}
					if(r.getRPr().getRFonts()!=null && style.getRPr().getRFonts()!=null){
						if(r.getRPr().getRFonts().equals(style.getRPr().getRFonts())){
							
						} else {
							isAccepted = false;
							break;
						}
					} else {
						// ignore
					}
					if(r.getRPr().getSz()!=null && style.getRPr().getSz()!=null){
						if(r.getRPr().getSz().equals(style.getRPr().getSz())){
							
						} else {
							isAccepted = false;
							break;
						}
					} else {
						// ignore
					}
				} else {
					if(isNull){
						continue;
					}
					isAccepted = false;
					break;
				}
			}
		}
		
		return isAccepted;
	}
	
	private static boolean isStyleTypeParagraph(P p, StyleDefinitionsPart stylePart){
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
			Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
			if(style!=null && style.getType()!=null && style.getType().equals("paragraph")){
				isAccepted = true;
			}
		} else if(p!=null && p.getContent()!=null && p.getContent().size()>0){
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
						Style style = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
						if(style != null && style.getType()!=null && !style.getType().equals("paragraph")){
							isAccepted = true;
							break;
						}
					}
				}
			}
		}
		
		return isAccepted;
	}
	
	private static boolean isBeginWithCaptionNum(P p, String pattern){
		boolean isAccepted = false;
		if(p!=null){
			String str = getPText(p);
			if(str.matches(pattern)){
				isAccepted = true;
			}
//			String alignmentString = alignment.value();
//			if(p.getPPr()!=null){
//				if(p.toString().matches(pattern)){
//					if(startStyle!=null){
//						if(isDeep && startStyle.getRPr()!=null){
//							if(p.getContent()!=null && p.getContent().size()>0){
//								for (int i = 0; i < p.getContent().size(); i++) {
//									Object o = p.getContent().get(i);
//									if(o instanceof R){
//										
//									}
//								}
//							}
//						} else {
//							if(p.getPPr().getRPr()!=null){
//								ParaRPr rPr = p.getPPr().getRPr();
//								
//								if(rPr.getB()!=null && rPr.getB().isVal()
//										&& startStyle.getRPr().getB()!=null
//										&& startStyle.getRPr().getB().isVal()
//										&& isBold){
//									isAccepted = true;
//								} else if(isBold){
//									isAccepted = false;
//								}
//								if(rPr.getI()!=null && rPr.getI().isVal()
//										&& startStyle.getRPr().getI()!=null
//										&& startStyle.getRPr().getI().isVal()
//										&& isItalic){
//									isAccepted = true;
//								} else if(isItalic){
//									isAccepted = false;
//								}
//								if(alignment!=null && isAcceptedFontAlignment(p, stylePart, alignmentString)){
//									isAccepted = true;
//								} else if(alignment!=null){
//									isAccepted = false;
//								}
//							}
//							if(startStyle.getStyleId()!=null && p.getPPr().getPStyle()!=null
//									&& startStyle.getStyleId().equals(p.getPPr().getPStyle().getVal())){
//								isAccepted = true;
//							}
//						}
//					}
//				} else {
//					// false
//				}
//			}
		}
		return isAccepted;
	}
	
	private static boolean isBeginWithCaptialNum(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart, String pattern, NumberFormat numFm){
		boolean isAccepted = false;
		if(p!=null){
			if(p!=null && p.getPPr()!=null){
				String str = getSectionCaptialNum(p);
				if(isSectionNum(str, pattern)){
					isAccepted = true;
				} else if(p.getPPr().getNumPr()!=null && p.getPPr().getNumPr().getNumId()!=null){
					NumPr numPr = p.getPPr().getNumPr();
					int iLvlVal = p.getPPr().getNumPr().getIlvl().getVal().intValue();
					NumId numId = p.getPPr().getNumPr().getNumId();
					if(iLvlVal>=0){
						NumberingDefinitionsPart ndp = documentPart.getNumberingDefinitionsPart();
						if(ndp!=null){
							try {
								if(isPStyleWithSectionNum(ndp, numId.getVal(), BigInteger.valueOf(iLvlVal), pattern, numFm)){
									isAccepted = true;
								}
							} catch (Exception e) {
								// TODO Auto-generated catch block
								LogUtils.log(e.getMessage());
							}
						}
					}
				} else if (p.getPPr().getPStyle()!=null){
					Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
					if(style!=null && style.getPPr()!=null && style.getPPr().getNumPr()!=null){
						NumPr numPr = style.getPPr().getNumPr();
						int iLvlVal = -1;
						NumId numId = null;
						if(numPr!=null){
							numId = numPr.getNumId();
							if(numId==null){
								NumPr numPrRek = rekursiveGetNumPr(style, stylePart);
								if(numPrRek!=null){
									numId = numPrRek.getNumId();
								}
							}
							Ilvl iLvl = numPr.getIlvl();
							if(iLvl!=null){
								iLvlVal = iLvl.getVal()!=null? iLvl.getVal().intValue():-1;
							} else {
								iLvlVal = 0;
							}
						}
						if(iLvlVal>=0){
							NumberingDefinitionsPart ndp = documentPart.getNumberingDefinitionsPart();
							if(ndp!=null){
								try {
									if(isPStyleWithSectionNum(ndp, numId.getVal(), BigInteger.valueOf(iLvlVal), pattern, numFm)){
										isAccepted = true;
									}
									
								} catch (Exception e) {
									// TODO Auto-generated catch block
									LogUtils.log(e.getMessage());
								}
								
							}
						}
					}
				} else if(p.getContent()!=null && p.getContent().size()>0){
					if(!isAllRInSameStyle(p, stylePart)){
						str = getSectionCaptialNum(p);
						if(isSectionNum(str, null)){
							isAccepted = true;
						}
					}
				}
			}
		} 
		return isAccepted;
	}
	
	private static NumPr rekursiveGetNumPr(Style style, StyleDefinitionsPart stylePart){
		if(style!=null){
			if(style.getPPr()!=null && style.getPPr().getNumPr()!=null && style.getPPr().getNumPr().getNumId()!=null){
				return style.getPPr().getNumPr();
			} else if(style.getBasedOn()!=null && style.getBasedOn().getVal()!=null){
				Style basedOnStyle = stylePart.getStyleById(style.getBasedOn().getVal());
				return rekursiveGetNumPr(basedOnStyle, stylePart);
			} else {
				return null;
			}
		} else {
			return null;
		}
	}
	private static boolean isHasPNum(P p, StyleDefinitionsPart stylePart){
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null){
			if(p.getPPr().getNumPr()!=null && p.getPPr().getNumPr().getNumId()!=null){
				isAccepted = true;
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style != null){
					if(style.getPPr()!=null && style.getPPr().getNumPr()!=null && style.getPPr().getNumPr().getNumId()!=null){
						isAccepted = true;
					}
					if(!isAccepted){
						NumPr numPr = rekursiveGetNumPr(style, stylePart);
						if(numPr!=null){
							isAccepted = true;
						}
					}
//					if(!isAccepted && style.getBasedOn()!=null){
//						Style styleBasedOn = stylePart.getStyleById(style.getBasedOn().getVal());
//						if(styleBasedOn!=null){
//							if(styleBasedOn.getPPr()!=null && styleBasedOn.getPPr().getNumPr()!=null && styleBasedOn.getPPr().getNumPr().getNumId()!=null){
//								isAccepted = true;
//							}
//						}
//					}
					if(!isAccepted && style.getLink()!=null){
						Style linkedStyle = stylePart.getStyleById(style.getLink().getVal());
						if(linkedStyle!=null){
							if(linkedStyle.getPPr()!=null && linkedStyle.getPPr().getNumPr()!=null && linkedStyle.getPPr().getNumPr().getNumId()!=null){
								isAccepted = true;
							}
						}
					}
				}
			} else {
				String numString = getSectionCaptialNum(p);
				if(numString!=null && numString.length()>0){
					isAccepted = true;
				}
			}
		}
		if(!isAccepted){
			boolean isIncludeSEQ = isIncludeSEQ(p);
			
			String pattern = "^(Figure|Fig.|Table|TABLE)\\s{1,1}\\d+..*";
			String pText = getPText(p);
//			if(pText.matches(pattern) && isIncludeSEQ){
			if(pText.matches(pattern)){
				isAccepted = true;
			}
//			pattern = "^Fîg.\\s{1,1}\\d+..*";
//			if(pText.matches(pattern) && isIncludeSEQ){
//				isAccepted = true;
//			}
//			pattern = "^Table\\s{1,1}\\d+..*";
//			if(pText.matches(pattern) && isIncludeSEQ){
//				isAccepted = true;
//			}
//			pattern = "^TABLE\\s{1,1}\\d+..*";
//			if(pText.matches(pattern) && isIncludeSEQ){
//				isAccepted = true;
//			}
			pattern = "^(([a-z]|\\d)\\)|([A-Z]|\\d(.\\d)+|\\d).|\\d(.\\d)+s)\\s{1,1}.*";
			if(pText.matches(pattern)){
				isAccepted = true;
			}
			pattern = "^(?=[MDCLXVI])M*(C[MD]|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})..*$";
			if(pText.matches(pattern)){
				isAccepted = true;
			}
		}
		return isAccepted;
	}
	
	public static boolean isCaption(P p, MainDocumentPart documentPart, StyleDefinitionsPart stylePart){
		boolean isCaption = false;
		
		// ACM Caption
		int fontSZ = 18; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		boolean isCaps = false;
		boolean isSmallCaps = false;
		String alignment = "center";
		String pattern = "^Figure\\s\\d+..*";
		NumberFormat numFm = NumberFormat.DECIMAL;
				
		if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& isAcceptedFontName(p, stylePart, fontName, false)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& isBeginWithCaptionNum(p, pattern)){
			isCaption = true;
		}
			
		// Elsevier Caption
		if(!isCaption){
			fontSZ = 16; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
//			alignment = "left";
			pattern = "^Fig.\\s{1,1}\\d+..*";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, true)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
//					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)){
				isCaption = true;
			}
		}		
		
		// IEEE Catption
		if(!isCaption){
			fontSZ = 16; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^Fig.\\s{1,1}\\d+..*";
			numFm = NumberFormat.DECIMAL;
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)){
				isCaption = true;
			}
		}
		
		// Springer Cation
		if(!isCaption){
			fontSZ = 18; 
			fontName = "Times"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^Fig.\\s{1,1}\\d+..*";
			numFm = NumberFormat.DECIMAL;
			Style specialStyle = new Style();
			specialStyle.setRPr(new RPr());
			specialStyle.getRPr().setB(new BooleanDefaultTrue());
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)
					&& !isAllRInSameStyle(p, stylePart)
					&& getTextWithSpecialStyle(p, stylePart, specialStyle)!=null){
				isCaption = true;
			}
		}
		return isCaption;
	}
	
	public static boolean isTableCaption(P p, StyleDefinitionsPart stylePart){
		boolean isTableCaption = false;
		
		// ACM Table Caption
		int fontSZ = 18; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		boolean isCaps = false;
		boolean isSmallCaps = false;
		String alignment = "center";
		String pattern = "^Table\\s\\d+..*";
		
		if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& isAcceptedFontName(p, stylePart, fontName, false)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& isBeginWithCaptionNum(p, pattern)){
			isTableCaption = true;
		}
		
		// Elsevier Table Caption
		if(!isTableCaption){
			fontSZ = 16; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			pattern = "^Table\\s{1,1}\\d+..*";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)){
				isTableCaption = true;
			}
		}
		
		// IEEE Table Caption
		if(!isTableCaption){
			fontSZ = 16; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "center";
			pattern = "^TABLE\\s{1,1}\\d+..*";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, true):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)){
				isTableCaption = true;
			}
		}
		
		// Springer Table Caption
		if(!isTableCaption){
			fontSZ = 18; 
			fontName = "Times"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^Table\\s{1,1}\\d+..*";
			Style specialStyle = new Style();
			specialStyle.setRPr(new RPr());
			specialStyle.getRPr().setB(new BooleanDefaultTrue());
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)
					&& !isAllRInSameStyle(p, stylePart)
					&& getTextWithSpecialStyle(p, stylePart, specialStyle)!=null){
				isTableCaption = true;
			}
		}
		
		return isTableCaption;
	}
	
	public static boolean adaptFirstWordWithStyle(P p, String text, RStyle rStyle){
		boolean result = false;
		
		
		
		return result;
	}
	
	public static boolean adaptFirstWord2Special(P p, String oldString, String expectedString, String charStyleId){
		boolean result = false;
		
		if(p!=null && expectedString!=null){
			boolean isReplaced = false;
			boolean isBeginSetStyle = false;
			Text lastText = null;
	EX_OUT:	for (int i = 0; i < p.getContent().size(); i++) {
				Object o = p.getContent().get(i);
				if(o instanceof R){
					for (int j = 0; j < ((R)o).getContent().size(); j++) {
						Object rO = ((R)o).getContent().get(j);
						if(rO instanceof JAXBElement<?> 
								&& ((JAXBElement) rO).getValue() instanceof Text
								&& !((JAXBElement) rO).getName().getLocalPart().equals("instrText")){
							Text text = (Text)(((JAXBElement) rO).getValue());
							if(text.getValue().equals(oldString)){	// the whole string should be replaced
								if(isReplaced){ // some text was cutted and replaced, so the text here should only be removed.
									text.setValue(null);
								} else {
									text.setValue(expectedString);
								}
								if(charStyleId!=null){
									((R) o).setRPr(Context.getWmlObjectFactory().createRPr());
									((R) o).getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
									((R) o).getRPr().getRStyle().setVal(charStyleId);
								} else {
									((R) o).setRPr(null);
								}
								result = true;
								break EX_OUT;
							} else if(text.getValue().indexOf(oldString)==0){	// only a part of string is need to replace
								if(isReplaced){
									text.setValue(text.getValue().replaceFirst(oldString, ""));
									result = true;
									break EX_OUT;
								} else {
									Text newText = Context.getWmlObjectFactory().createText();
									newText.setValue(expectedString);
									R newR = Context.getWmlObjectFactory().createR();
									if(charStyleId!=null){
										newR.setRPr(Context.getWmlObjectFactory().createRPr());
										newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
										newR.getRPr().getRStyle().setVal(charStyleId);
									} else {
										newR.setRPr(null);
									}
									
									newR.getContent().add(0, newText);
									lastText = newText;
									p.getContent().add(0, newR);
									text.setValue(text.getValue().replaceFirst(oldString, ""));
									isBeginSetStyle = true;
								}
								
							} else if(oldString.indexOf(text.getValue())==0 && j<oldString.length()){
								if(!isReplaced){
									oldString = oldString.substring(text.getValue().length(), oldString.length());
									text.setValue(expectedString);
									if(charStyleId!=null){
										((R) o).setRPr(Context.getWmlObjectFactory().createRPr());
										((R) o).getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
										((R) o).getRPr().getRStyle().setVal(charStyleId);
									} else {
										((R) o).setRPr(null);
									}
									isReplaced = true;
								}
							} else if(text.getValue().replaceAll(" ", "").equals(expectedString)){
								if(charStyleId!=null){
									((R) o).setRPr(Context.getWmlObjectFactory().createRPr());
									((R) o).getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
									((R) o).getRPr().getRStyle().setVal(charStyleId);
								} else {
									((R) o).setRPr(null);
								}
								lastText = text;
								isBeginSetStyle = true;
							} else if(isBeginSetStyle){
								if(text.getValue().indexOf(".")>=0 && text.getValue().indexOf(".")<8){
									String str = text.getValue().substring(0, text.getValue().indexOf(".")+1);
									String str2 = text.getValue().substring(text.getValue().indexOf(".")+1, text.getValue().length());
									if(lastText!=null){
										lastText.setValue(lastText.getValue()+str);
										text.setValue(str2);
										text.setSpace("preserve");
									} 
									result = true;
									break EX_OUT;
								} else {
									if(charStyleId!=null){
										((R) o).setRPr(Context.getWmlObjectFactory().createRPr());
										((R) o).getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
										((R) o).getRPr().getRStyle().setVal(charStyleId);
									} else {
										((R) o).setRPr(null);
									}
									lastText = text;
								}
							}
						}
					}
				}
			}
		}
		
		return result;
	}
	
	public static boolean isEqualSpecialText(P p){
		boolean isAccepted = false;
		String text = getPText(p);
		text = text.toLowerCase();
		if("reference".equals(text)){
			isAccepted = true;
		} else if("references".equals(text)){
			isAccepted = true;
		} else if("acknowledgment".equals(text)){
			isAccepted = true;
		} else if("acknowledgments".equals(text)){
			isAccepted = true;
		}
		return isAccepted;
	}
	
	public static boolean isAcknowledgReferenceHeader(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart){
		boolean isAcknowledgReferenceHeader = false;
		
		String text = getPText(p);
		boolean isSpecial = false;
		if(text.replaceAll(" ", "").toLowerCase().equals("Acknowledgment".toLowerCase()) || text.replaceAll(" ", "").toLowerCase().equals("References".toLowerCase())){
			isSpecial = true;
		}
		// ACM ReferenceHeader
		int fontSZ = 24; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		boolean isCaps = true;
		boolean isSmallCaps = false;
		String alignment = "left";
		String numberPattern = "%1.";
		NumberFormat numFm = NumberFormat.DECIMAL;
		
		if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& isAcceptedFontName(p, stylePart, fontName, false)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, true):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& isHasPNum(p, stylePart)
				&& isBeginWithCaptialNum(p, stylePart, documentPart, numberPattern, numFm)
				&& isEqualSpecialText(p)
				){
			isAcknowledgReferenceHeader = true;
		}
		
		// Elsevier ReferenceHeader
		if(!isAcknowledgReferenceHeader){
			fontSZ = 20; 
			fontName = "Times New Roman"; 
			isBold = true; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& !isHasPNum(p, stylePart)
					&& isEqualSpecialText(p)
					){
				isAcknowledgReferenceHeader = true;
			}
		}
		
		// IEEE ReferenceHeader
		if(!isAcknowledgReferenceHeader){
			fontSZ = 20; 
			fontName = "Times New Roman";
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = true;
			alignment = "center";
			if((isAcceptedFontSZ(p, fontSZ, stylePart, false)||isSpecial && sourceFormat.equals(SOURCEFORMAT_IEEE))
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& !isHasPNum(p, stylePart)
					&& isEqualSpecialText(p)
					){
				isAcknowledgReferenceHeader = true;
			}
		}
		
		// Springer ReferenceHeader
		if(!isAcknowledgReferenceHeader){
			fontSZ = 24;
			fontName = "Times"; 
			isBold = true; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& !isHasPNum(p, stylePart)
					&& isEqualSpecialText(p)
					){
				isAcknowledgReferenceHeader = true;
			}
		}
		
		return isAcknowledgReferenceHeader;
	}
	
	private static boolean isSpaceAccepted(P p, Style specialStyle, StyleDefinitionsPart stylePart, MainDocumentPart documentPart, boolean isStrict){
		boolean isAccepted = false;
		
		if(specialStyle==null){
			return isAccepted;
		}
		Ind chkInd = null;
		if(specialStyle!=null && specialStyle.getPPr()!=null && specialStyle.getPPr().getInd()==null){
			return isAccepted;
		} else {
			chkInd = specialStyle.getPPr().getInd();
		}
		
		if(p!=null && p.getPPr()!=null){
			Ind ind = null;
			if(p.getPPr().getInd()!=null){
				ind = p.getPPr().getInd();
			} else if(p.getPPr().getPStyle()!=null && p.getPPr().getPStyle().getVal()!=null){
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(style!=null && style.getPPr()!=null && style.getPPr().getInd()!=null){
					ind = style.getPPr().getInd();
				} else if(style.getPPr()!=null && style.getPPr().getNumPr()!=null){
					try {
						BigInteger numId = style.getPPr().getNumPr().getNumId().getVal();
						NumberingDefinitionsPart numberingPart = documentPart.getNumberingDefinitionsPart();
						Numbering numbering = numberingPart.getContents();
						if(numbering!=null){
							List<Num> numList = numbering.getNum();
							BigInteger abstractNumId = null;
							for(int numIndex=0; numIndex<numList.size(); numIndex++) {
								if(numList.get(numIndex).getNumId().intValue() == numId.intValue()) {
									abstractNumId = numList.get(numIndex).getAbstractNumId().getVal();
									break;
								}
							}
							if(abstractNumId!=null){
								List<AbstractNum> abstractNumList = numbering.getAbstractNum();
					NUMINDEX:	for(int aNumIndex=0; aNumIndex<abstractNumList.size(); aNumIndex++) {
									if(abstractNumList.get(aNumIndex).getAbstractNumId().intValue() == abstractNumId.intValue()){
										List<Lvl> lvlList = abstractNumList.get(aNumIndex).getLvl();
										for(int lvlIndex=0; lvlIndex<lvlList.size(); lvlIndex++) {
											if(lvlList.get(lvlIndex).getPPr()!=null && lvlList.get(lvlIndex).getPPr().getInd()!=null){
												ind = lvlList.get(lvlIndex).getPPr().getInd();
												break NUMINDEX;
											}
										}
									}
								}
							}
						}
					} catch (Exception e) {
						// TODO: handle exception
						LogUtils.log(e.getMessage());
					}
				}
			}
			if(ind!=null){
				if(chkInd.getFirstLine()!=null){
					if(ind.getFirstLine()!=null && ind.getFirstLine().intValue() == chkInd.getFirstLine().intValue()){
						isAccepted = true;
					} else {
						isAccepted = false;
						return isAccepted;
					}
				}
				if(chkInd.getHanging()!=null){
					if(ind.getHanging()!=null && ind.getHanging().intValue() == chkInd.getHanging().intValue()){
						isAccepted = true;
					} else {
						isAccepted = false;
						return isAccepted;
					}
				}
				if(chkInd.getLeft()!=null){
					if(ind.getLeft()!=null && ind.getLeft().intValue() == chkInd.getLeft().intValue()){
						isAccepted = true;
					} else {
						isAccepted = false;
						return isAccepted;
					}
				}
				if(!isStrict){
					isAccepted = true;
				}
			}
		}
		return isAccepted;
	}
	
	public static boolean isReferenceText(P p, StyleDefinitionsPart stylePart, MainDocumentPart documentPart){
		boolean isReferenceText = false;
		
		// ACM Reference text
		int fontSZ = 18; 
		String fontName = "Times New Roman"; 
		boolean isBold = true; 
		boolean isItalic = false;
		boolean isCaps = false;
		boolean isSmallCaps = false;
		String alignment = "center";
		String pattern = "^\\[\\d\\]\\s{1,1}.*";
		Style specialStyle = new Style();
		specialStyle.setPPr(new PPr());;
		specialStyle.getPPr().setInd(new Ind());
		specialStyle.getPPr().getInd().setFirstLine(null);
		specialStyle.getPPr().getInd().setHanging(new BigInteger("360"));
		specialStyle.getPPr().getInd().setLeft(new BigInteger("360"));
		
		if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
				&& isAcceptedFontName(p, stylePart, fontName, false)
				&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
				&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
				&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
				&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
				&& isAcceptedFontAlignment(p, stylePart, alignment)
				&& isBeginWithCaptionNum(p, pattern)
				&& isSpaceAccepted(p, specialStyle, stylePart, documentPart, false)
				){
			isReferenceText = true;
		}
				
		// Elsevier Reference Text
		if(!isReferenceText){
			fontSZ = 16; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "left";
			specialStyle = new Style();
			specialStyle.setPPr(new PPr());;
			specialStyle.getPPr().setInd(new Ind());
			specialStyle.getPPr().getInd().setFirstLine(null);
			specialStyle.getPPr().getInd().setHanging(new BigInteger("360"));
			specialStyle.getPPr().getInd().setLeft(new BigInteger("360"));
			
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isSpaceAccepted(p, specialStyle, stylePart, documentPart, true)
					){
				isReferenceText = true;
			}
		}
				
		// IEEE Reference Text
		if(!isReferenceText){
			fontSZ = 16; 
			fontName = "Times New Roman"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^\\[\\d\\]\\s{1,1}.*";
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)
					&& isSpaceAccepted(p, specialStyle, stylePart, documentPart, false)
					){
				isReferenceText = true;
			}
		}
				
		// Springer Reference Text
		if(!isReferenceText){
			fontSZ = 18; 
			fontName = "Times"; 
			isBold = false; 
			isItalic = false;
			isCaps = false;
			isSmallCaps = false;
			alignment = "both";
			pattern = "^\\[\\d\\]\\s{1,1}.*";
			
			if(isAcceptedFontSZ(p, fontSZ, stylePart, false)
					&& isAcceptedFontName(p, stylePart, fontName, false)
					&& (isBold?isAcceptedFontBold(p, stylePart, false):true)
					&& (isItalic?isAcceptedFontItalic(p, stylePart, false):true)
					&& (isCaps?isAcceptedFontCaps(p, stylePart, false):true)
					&& (isSmallCaps?isAcceptedFontSmallCaps(p, stylePart, false):true)
					&& isAcceptedFontAlignment(p, stylePart, alignment)
					&& isBeginWithCaptionNum(p, pattern)
					&& !isAllRInSameStyle(p, stylePart)
//					&& getTextWithSpecialStyle(p, stylePart, specialStyle)!=null
					&& isSpaceAccepted(p, specialStyle, stylePart, documentPart, false)
					){
				isReferenceText = true;
			}
		}
		
		return isReferenceText;
	}
	
	public static void fixImgAlignment(MainDocumentPart documentPart) {
		try {
			ClassFinder classFiner = new ClassFinder(Drawing.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			for(Object drawingObj : classFiner.results){
				Drawing drawing = (Drawing) drawingObj;
				if(drawing.getParent()!=null && drawing.getParent() instanceof R){
					R r = (R) drawing.getParent();
					if(r.getParent()!=null && r.getParent() instanceof P){
						P p = (P) r.getParent();
						if(p.getPPr()!=null){
							Jc jc = new Jc();
							jc.setVal(JcEnumeration.CENTER);
							p.getPPr().setJc(jc);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
	}
	
	public static void fixImgWidth(MainDocumentPart documentPart){
		try {
			ClassFinder classFiner = new ClassFinder(Inline.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			for(Object inlineObj : classFiner.results){
				Inline inline = (Inline) inlineObj;
				CTPositiveSize2D extent = inline.getExtent();
				long cx = extent.getCx();
				long cy = extent.getCy();
				if(cx>12240){
					extent.setCx(3*914400);
					int calY = (int) (cy*3/cx*914400);
					extent.setCy(calY);
//					LogUtils.log("image size from ("+cx+", "+ cy+") to ("+extent.getCx()+", "+ extent.getCy()+")");
					if(inline.getGraphic()!=null && inline.getGraphic().getGraphicData()!=null 
							&& inline.getGraphic().getGraphicData().getPic()!=null
							&& inline.getGraphic().getGraphicData().getPic().getSpPr()!=null
							&& inline.getGraphic().getGraphicData().getPic().getSpPr().getXfrm()!=null
							&& inline.getGraphic().getGraphicData().getPic().getSpPr().getXfrm().getExt()!=null){
						inline.getGraphic().getGraphicData().getPic().getSpPr().getXfrm().getExt().setCx(extent.getCx());
						inline.getGraphic().getGraphicData().getPic().getSpPr().getXfrm().getExt().setCy(extent.getCy());
//						LogUtils.logShort("image size changed\n");
					}
				} 
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
	}
	public static boolean isImage(P p){
		boolean isImage = false;
		
		if(p!=null){
			List<Object> pContent = p.getParagraphContent();
	EXITP:	for (int i = 0; i < pContent.size(); i++) {
				Object pO = pContent.get(i);
				if(pO instanceof R){
					R r = (R)pO;
					List<Object> rContent = r.getRunContent();
					for (int j = 0; j < rContent.size(); j++) {
						Object rO = rContent.get(j);
						if(rO instanceof JAXBElement<?> && ((JAXBElement) rO).getValue() instanceof Drawing){
							isImage = true;
							break EXITP;
						}
					}
				}
			}
		}
		
		return isImage;
	}
	
	public static void adaptPageLayout(SectPr sectPr, String colNumber, String colSpacing,
			String pgSZW, String pgSZH, String pgSZCode,
			String pgMarT, String pgMarB, String pgMarL, String pgMarR, String pgMarHeader, String pgMarFooter, String pgMarGutter){
		if(pgSZCode!=null){
			PgSz pgSZ = Context.getWmlObjectFactory().createSectPrPgSz();
			if(pgSZH!=null){
				pgSZ.setH(new BigInteger(pgSZH));
			}
			if(pgSZW!=null) {
				pgSZ.setW(new BigInteger(pgSZW));
			}
			if(!pgSZCode.equals("-1")) {
				pgSZ.setCode(new BigInteger(pgSZCode));
				sectPr.setPgSz(pgSZ);
			}
		}
		
		if(true){
			PgMar pgMar = Context.getWmlObjectFactory().createSectPrPgMar();
			if(pgMarB!=null){
				pgMar.setBottom(new BigInteger(pgMarB));
			}
			if(pgMarT!=null){
				pgMar.setTop(new BigInteger(pgMarT));
			}
			if(pgMarL!=null){
				pgMar.setLeft(new BigInteger(pgMarL));
			}
			if(pgMarR!=null){
				pgMar.setRight(new BigInteger(pgMarR));
			}
			if(pgMarHeader!=null){
				pgMar.setHeader(new BigInteger(pgMarHeader));
			}
			if(pgMarFooter!=null){
				pgMar.setFooter(new BigInteger(pgMarFooter));
			}
			if(pgMarGutter!=null){
				pgMar.setGutter(new BigInteger(pgMarGutter));
			}
			if(pgMarT!=null && pgMarB!=null || pgMarL!=null && pgMarR!=null){
				sectPr.setPgMar(pgMar);
			}
		}
		
		if(colNumber!=null && colSpacing!=null){
			CTColumns cols = Context.getWmlObjectFactory().createCTColumns();
			cols.setNum(new BigInteger(colNumber));
			cols.setSpace(new BigInteger(colSpacing));
			sectPr.setCols(cols);
		}
		
		sectPr.setType(Context.getWmlObjectFactory().createSectPrType());
		sectPr.getType().setVal("continuous");
	}
	public static void fixPageLayout(SectPr sectPr, String colNumber, String colSpacing){
		fixPageLayout(sectPr, colNumber, colSpacing, true);
	}
	public static void fixPageLayout(SectPr sectPr, String colNumber, String colSpacing, boolean isColSpacingOnly){
		if(!isColSpacingOnly){
			PgSz pgSZ = sectPr.getPgSz();
			if(pgSZ!=null){
				
			} else {
				pgSZ = new PgSz();
				sectPr.setPgSz(pgSZ);
			}
			pgSZ.setCode(new BigInteger("1"));
			pgSZ.setW(new BigInteger("12240"));
			pgSZ.setH(new BigInteger("15840"));
			
			PgMar pgMar = sectPr.getPgMar();
			if(pgMar!=null){
				
			} else {
				pgMar = new PgMar();
				sectPr.setPgMar(pgMar);
			}
			pgMar.setTop(new BigInteger("1080"));
			pgMar.setRight(new BigInteger("1080"));
			pgMar.setBottom(new BigInteger("1440"));
			pgMar.setLeft(new BigInteger("1080"));
			pgMar.setHeader(new BigInteger("720"));
			pgMar.setFooter(new BigInteger("720"));
			pgMar.setGutter(new BigInteger("0"));
			
		}
		
		CTColumns cols = sectPr.getCols();
		if(cols!=null){
			
		} else {
			cols = new CTColumns();
			sectPr.setCols(cols);
		}
		cols.setNum(new BigInteger(colNumber));
		cols.setSpace(new BigInteger(colSpacing));
		
		/*
		 * keep section together
		 */
		Type type = sectPr.getType();
		if(type!=null){
			
		} else {
			type = new Type();
			sectPr.setType(type);
		}
		type.setVal("continuous");
	}
	
	public static boolean isEmail(P p){
		boolean isEmail = false;
		
//		String regex = "^[A-Za-z0-9\\._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,6}$";
//		isEmail = p.toString().matches(regex);
		String emailText = getPText(p);
		Pattern VALID_EMAIL_ADDRESS_REGEX = Pattern.compile("^[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,6}$", Pattern.CASE_INSENSITIVE);
		Matcher matcher = VALID_EMAIL_ADDRESS_REGEX .matcher(emailText);
		isEmail = matcher.find();
		
		return isEmail;
	}
	
	public static void removePSetting(PPr pPr){
		if(pPr!=null){
			pPr.setInd(null);
			pPr.setJc(null);
			pPr.setNumPr(null);
			pPr.setSectPr(null);
			pPr.setTabs(null);
			pPr.setTextAlignment(null);
			
			pPr.setRPr(new ParaRPr());
			
			pPr.getRPr().setLang(new CTLanguage());
			pPr.getRPr().getLang().setVal("en-US");
		}
	}
	
	public static void removeAllProperties(P p){
		if(p!=null){
			if(p.getPPr()!=null){
				p.setPPr(null);
			}
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(null);
				}
			}
		}
	}
	
	public static void createAppendingLine(P p, Style style, int colNum){
		String colSpace = "720";
		String pgSzWidth = "12240";
		String pgSzHeight = "15840";
		String pgSzCode = "1";
		String pgMarginTop = "1080";
		String pgMarginBottom = "1440";
		String pgMarginLeft = "1080";
		String pgMarginRight = "1080";
		String pgMarginGutter = "0";
		String pgMarginFooter = "720";
		String pgMarginHeader = "720";
		
		SectPr sectPr = Context.getWmlObjectFactory().createSectPr();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setSectPr(sectPr);
		if(style.getPPr()!=null && style.getPPr().getSectPr()!=null){
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(colNum));
			sectPr.getCols().setSpace(style.getPPr().getSectPr().getCols().getSpace());
		} else {
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(colNum));
			sectPr.getCols().setSpace(new BigInteger(colSpace));
		}
		
		if(style.getPPr()!=null && style.getPPr().getSectPr()!=null){
			if(style.getPPr().getSectPr().getPgSz()!=null){
				sectPr.setPgSz(style.getPPr().getSectPr().getPgSz());
			} else {
				sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
			}
			if(sectPr.getPgSz()!=null){
				if(sectPr.getPgSz().getW()!=null){
					
				} else {
					sectPr.getPgSz().setW(new BigInteger(pgSzWidth));
				}
				if(sectPr.getPgSz().getH()!=null){
					
				} else {
					sectPr.getPgSz().setH(new BigInteger(pgSzHeight));
				}
			} else {
				sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
				sectPr.getPgSz().setW(new BigInteger(pgSzWidth));
				sectPr.getPgSz().setH(new BigInteger(pgSzHeight));
			}
			
			if(style.getPPr().getSectPr().getPgMar()!=null){
				sectPr.setPgMar(style.getPPr().getSectPr().getPgMar());
			} else {
				sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
			}
			if(sectPr.getPgMar()!=null){
				PgMar margin = sectPr.getPgMar();
				if(margin.getTop()!=null){
					
				} else {
					margin.setTop(new BigInteger(pgMarginTop));
				}
				if(margin.getBottom()!=null){
					
				} else {
					margin.setBottom(new BigInteger(pgMarginBottom));
				}
				if(margin.getLeft()!=null){
					
				} else {
					margin.setLeft(new BigInteger(pgMarginLeft));
				}
				if(margin.getRight()!=null){
					
				} else {
					margin.setRight(new BigInteger(pgMarginRight));
				}
				if(margin.getHeader()!=null){
					
				} else {
					margin.setHeader(new BigInteger(pgMarginHeader));
				}
				if(margin.getFooter()!=null){
					
				} else {
					margin.setFooter(new BigInteger(pgMarginFooter));
				}
				if(margin.getGutter()!=null){
					
				} else {
					margin.setGutter(new BigInteger(pgMarginGutter));
				}
			} else {
				sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				PgMar margin = sectPr.getPgMar();
				margin.setTop(new BigInteger(pgMarginTop));
				margin.setBottom(new BigInteger(pgMarginBottom));
				margin.setLeft(new BigInteger(pgMarginLeft));
				margin.setRight(new BigInteger(pgMarginRight));
				margin.setHeader(new BigInteger(pgMarginHeader));
				margin.setFooter(new BigInteger(pgMarginFooter));
				margin.setGutter(new BigInteger(pgMarginGutter));
			}
			/*
			 * keep section together
			 */
			Type type = sectPr.getType();
			if(type!=null){
				
			} else {
				type = new Type();
				sectPr.setType(type);
			}
			type.setVal("continuous");
		}
	}
	

	public static void adaptPage(SectPr sectPr, Style style, int colNum){
		String colSpace = "720";
		String pgSzWidth = "12240";
		String pgSzHeight = "15840";
		String pgSzCode = "1";
		String pgMarginTop = "1080";
		String pgMarginBottom = "1440";
		String pgMarginLeft = "1080";
		String pgMarginRight = "1080";
		String pgMarginGutter = "0";
		String pgMarginFooter = "720";
		String pgMarginHeader = "720";
		
		if(style.getPPr()!=null && style.getPPr().getSectPr()!=null){
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(colNum));
			sectPr.getCols().setSpace(style.getPPr().getSectPr().getCols().getSpace());
		} else {
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(colNum));
			sectPr.getCols().setSpace(new BigInteger(colSpace));
		}
		
		if(style.getPPr()!=null && style.getPPr().getSectPr()!=null){
			if(style.getPPr().getSectPr().getPgSz()!=null){
				sectPr.setPgSz(style.getPPr().getSectPr().getPgSz());
			} else {
				sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
			}
			if(sectPr.getPgSz()!=null){
				if(sectPr.getPgSz().getW()!=null){
					
				} else {
					sectPr.getPgSz().setW(new BigInteger(pgSzWidth));
				}
				if(sectPr.getPgSz().getH()!=null){
					
				} else {
					sectPr.getPgSz().setH(new BigInteger(pgSzHeight));
				}
			} else {
				sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
				sectPr.getPgSz().setW(new BigInteger(pgSzWidth));
				sectPr.getPgSz().setH(new BigInteger(pgSzHeight));
			}
			
			if(style.getPPr().getSectPr().getPgMar()!=null){
				sectPr.setPgMar(style.getPPr().getSectPr().getPgMar());
			} else {
				sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
			}
			if(sectPr.getPgMar()!=null){
				PgMar margin = sectPr.getPgMar();
				if(margin.getTop()!=null){
					
				} else {
					margin.setTop(new BigInteger(pgMarginTop));
				}
				if(margin.getBottom()!=null){
					
				} else {
					margin.setBottom(new BigInteger(pgMarginBottom));
				}
				if(margin.getLeft()!=null){
					
				} else {
					margin.setLeft(new BigInteger(pgMarginLeft));
				}
				if(margin.getRight()!=null){
					
				} else {
					margin.setRight(new BigInteger(pgMarginRight));
				}
				if(margin.getHeader()!=null){
					
				} else {
					margin.setHeader(new BigInteger(pgMarginHeader));
				}
				if(margin.getFooter()!=null){
					
				} else {
					margin.setFooter(new BigInteger(pgMarginFooter));
				}
				if(margin.getGutter()!=null){
					
				} else {
					margin.setGutter(new BigInteger(pgMarginGutter));
				}
			} else {
				sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
				PgMar margin = sectPr.getPgMar();
				margin.setTop(new BigInteger(pgMarginTop));
				margin.setBottom(new BigInteger(pgMarginBottom));
				margin.setLeft(new BigInteger(pgMarginLeft));
				margin.setRight(new BigInteger(pgMarginRight));
				margin.setHeader(new BigInteger(pgMarginHeader));
				margin.setFooter(new BigInteger(pgMarginFooter));
				margin.setGutter(new BigInteger(pgMarginGutter));
			}
			/*
			 * keep section together
			 */
			Type type = sectPr.getType();
			if(type!=null){
				
			} else {
				type = new Type();
				sectPr.setType(type);
			}
			type.setVal("continuous");
		}
	}
	
	public static boolean isBeginWithListNumber(P p, NumberingDefinitionsPart ndp, StyleDefinitionsPart stylePart){
		boolean isBeginWithChapterNumber = false;
		
		if(p!=null){
			if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
				if(p.getPPr().getNumPr()!=null) {
					isBeginWithChapterNumber = true;
				} else if(style!=null && style.getPPr()!=null && style.getPPr().getNumPr()!=null){
					isBeginWithChapterNumber = true;
				} else {
					String pText = OOXMLConvertTool.getPText(p);
					if(pText!=null){
						String pattern = "^\\d+(\\.\\d+)*([ ]{1,1})(.*)";
						if(isSectionNum(pText, pattern)){
							isBeginWithChapterNumber = true;
						}
					}
				}
			}
		}
		
		
		return isBeginWithChapterNumber;
	}
	
	public static void removeStringFromPLostRStyle(P p, String str){
		removeStringFromPLostRStyle(p, str, true);
	}
	
	public static void removeStringFromPLostRStyle(P p, String str, boolean isNeedRpr){
		int i = 0;
		if(str==null){
			return;
		}
		String pText = getPText(p);
		if(pText.indexOf(str)<0){
			return;
		} else {
			pText = pText.replaceFirst(str, "");
			RPr rPr = null;
			if(p.getContent()!=null && isNeedRpr){
				while(i<p.getContent().size()){
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null){
							rPr = r.getRPr();
							break;
						}
					}
					i++;
				}
			}
			p.getContent().clear();
			R newR = Context.getWmlObjectFactory().createR();
			Text text = Context.getWmlObjectFactory().createText();
			text.setValue(pText);
			newR.setRPr(rPr);
			newR.getContent().add(text);
			p.getContent().add(newR);
		}
	}
	
	public static String getStringOf(P p, String regex){
		String str = "";
		
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(getPText(p));
		if(matcher.find()){
			str = matcher.group(0);
		}
		
		return str;
	}
	
	public static void removeSectionNumStringOfHeading(P p){
		
		String[] regex = {"^\\d+(\\.\\d+)*\\.{0,1}",
			"^[a-z][\\)]{1,1}?",
			"^[A-Z][\\)]{1,1}?",
			"^\\d+[\\)]{1,1}?",
			"^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$"};
		for (int i = 0; i < regex.length; i++) {
			String sectionNumString = getStringOf(p, regex[i]);
			if(sectionNumString!=null && sectionNumString.length()>0){
				removeStringFromPLostRStyle(p, sectionNumString);
				return;
			}
		}
		if(p!=null && p.getPPr()!=null){
			p.getPPr().setNumPr(null);
		}
	}
	
	public static void removeNumStringFromString(P p, String regex){
		String numStr = getStringOf(p, regex);
		if(numStr!=null && numStr.length()>0){
			removeStringFromPLostRStyle(p, numStr);
		}
	}
	
	public static Object deepCloneNotSerializable(Object object) {
		Object o = null;
		try {
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			ObjectOutputStream oos = new ObjectOutputStream(baos);
			oos.writeObject(object);
			ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
			ObjectInputStream ois = new ObjectInputStream(bais);
			o = ois.readObject();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return o;
	}
	
	public static P cloneP(P p){
		P newP = new P();
		
		newP.setParaId(p.getParaId());
		newP.setParent(p.getParent());
		newP.setPPr(p.getPPr());
		newP.setRsidDel(p.getRsidDel());
		newP.setRsidP(p.getRsidP());
		newP.setRsidR(p.getRsidR());
		newP.setRsidRDefault(p.getRsidRDefault());
		newP.setRsidRPr(p.getRsidRPr());
		newP.setTextId(p.getTextId());
		newP.getContent().addAll(p.getContent());
		
		return newP;
	}
	
	public static Style getFirstRStyle(P p){
		Style rStyle = null;
		
		if(p!=null) {
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				ROOT:for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					String text = null;
					if(o instanceof R){
						R r = (R)o;
						text = getRText(r);
						if(text!=null && text.length()>0){
							if(r.getRPr()!=null){
								rStyle = new Style();
								rStyle.setRPr(r.getRPr());
							} else {
								
							}
							
							break ROOT;
						}
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R sfR = (R)sfO;
										text = getRText(sfR);
										if(text!=null && text.length()>0){
											if(sfR.getRPr()!=null){
												rStyle = new Style();
												rStyle.setRPr(sfR.getRPr());
											} else {
												
											}
											
											break ROOT;
										}
									}
								}
							}
						}
					} else {
//						System.out.println(o.getClass().getName());
					}
				}
			}
		}
		
		return rStyle;
	}
	
	public static P getSameRJudgeWithRRp(P p, StyleDefinitionsPart stylePart, boolean isDelFoundsFromP){
		P newP = null;
		
		if(p!=null){
			List<Object> pContents = p.getContent();
			newP = Context.getWmlObjectFactory().createP();
			
			RPr foundRPr = null;
			boolean isInitialized = false;
			boolean isRprChked = false;
			int i=0;
			for(;i<pContents.size();i++){
				Object o = pContents.get(i);
				if(o instanceof R){
					R r = (R)o;
					String rText = getRText(r);
					if(rText!=null && rText.replaceAll(" ", "").length()>0){
						if(newP.getPPr()==null){
							newP.setPPr(Context.getWmlObjectFactory().createPPr());
						}
						
						if(r.getRPr()!=null && !isInitialized){
							if(foundRPr==null){
								foundRPr = r.getRPr();
								isRprChked = true;
							} else if(!isInitialized){
								if(r.getRPr().getB()!=null && foundRPr.getB()!=null){
									if(r.getRPr().getB().isVal() && foundRPr.getB().isVal()){
										
									} else{
										break;
									}
								} else if(r.getRPr().getB()==null && foundRPr.getB()==null){
									
								} else {
									break;
								}
								if(r.getRPr().getI()!=null && foundRPr.getI()!=null){
									if(r.getRPr().getI().isVal() && foundRPr.getI().isVal()){
										
									} else{
										break;
									}
								} else if(r.getRPr().getI()==null && foundRPr.getI()==null){
									
								} else {
									break;
								}
								if(r.getRPr().getSmallCaps()!=null && foundRPr.getSmallCaps()!=null){
									if(r.getRPr().getSmallCaps().isVal() && foundRPr.getSmallCaps().isVal()){
										
									} else{
										break;
									}
								} else if(r.getRPr().getSmallCaps()==null && foundRPr.getSmallCaps()==null){
									
								} else {
									break;
								}
								if(r.getRPr().getCaps()!=null && foundRPr.getCaps()!=null){
									if(r.getRPr().getCaps().isVal() && foundRPr.getCaps().isVal()){
										
									} else{
										break;
									}
								} else if(r.getRPr().getCaps()==null && foundRPr.getCaps()==null){
									
								} else {
									break;
								}
								if(r.getRPr().getSz()!=null && foundRPr.getSz()!=null){
									if(r.getRPr().getSz().getVal()!=null && foundRPr.getSz().getVal()!=null
											&& r.getRPr().getSz().getVal().intValue() == foundRPr.getSz().getVal().intValue()){
										
									} else{
										break;
									}
								} else if(r.getRPr().getSz()==null && foundRPr.getSz()==null){
									
								} else {
									break;
								}
								if(r.getRPr().getRFonts()!=null && foundRPr.getRFonts()!=null){
									if(r.getRPr().getRFonts().getAscii().equals(foundRPr.getRFonts().getAscii())){
										
									} else {
										break;
									}
								} else if(r.getRPr().getRFonts()==null && foundRPr.getRFonts()==null) {
									
								} else {
									break;
								}
								if(r.getRPr().getRStyle()!=null && foundRPr.getRStyle()!=null){
									if(r.getRPr().getRStyle().getVal().equals(foundRPr.getRStyle().getVal())){
										
									} else {
										break;
									}
								} else if(r.getRPr().getRStyle()==null && foundRPr.getRStyle()==null){
									
								} else {
									break;
								}
							} else {
								break;
							}
						} else {
							if(!isInitialized && !isRprChked) {
								isInitialized = true;
							} else {
								break;
							}
						}
						
						newP.getContent().add(r);
						if(isDelFoundsFromP){
							p.getContent().remove(i);
						}
					}
				}
			}
		}
		
		return newP;
	}
	
	public static void replaceFirstSameGroupRStyle(P p, Style style, StyleDefinitionsPart stylePart, String oldString, String addString){

		List<Object> pContents = p.getContent();
		RPr rPr = null;
		int index = -1;
		for (int i = 0; i < pContents.size(); i++) {
			Object o = pContents.get(i);
			if(o instanceof R){
				R r = (R)o;
				String rText = getRText(r);
				if(rText==null || rText!=null && rText.replaceAll(" ", "").length()==0){
					p.getContent().remove(o);
					continue;
				}
				if(rPr==null && index<0){
					rPr = r.getRPr();
					index = i;
				} else {
					Style rStyle = (r.getRPr()!=null && r.getRPr().getRStyle()!=null)?stylePart.getStyleById(r.getRPr().getRStyle().getVal()):null;
					Style firstRStyle = (rPr!=null && rPr.getRStyle()!=null)?stylePart.getStyleById(rPr.getRStyle().getVal()):null;
					
					if(r.getRPr()!=null && rPr!=null){
						if(rPr.getB()!=null && r.getRPr().getB()!=null){
							if(rPr.getB().isVal() && r.getRPr().getB().isVal()){
								
							} else {
								index = i-1;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getB()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getB()!=null 
										&& firstRStyle.getRPr().getB().isVal()){
									chk1 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(r.getRPr().getB()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getB()!=null
										&& rStyle.getRPr().getB().isVal()){
									chk2 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								index = i-1;
								break;
							}
						}
						if(rPr.getI()!=null && r.getRPr().getI()!=null){
							if(rPr.getI().isVal() && r.getRPr().getI().isVal()){
								
							} else {
								index = i-1;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getI()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getI()!=null 
										&& firstRStyle.getRPr().getI().isVal()){
									chk1 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(r.getRPr()==null || r.getRPr()!=null && r.getRPr().getI()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getI()!=null
										&& rStyle.getRPr().getI().isVal()){
									chk2 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								index = i-1;
								break;
							}
						}
						if(rPr.getCaps()!=null && r.getRPr().getCaps()!=null){
							if(rPr.getCaps().isVal() && r.getRPr().getCaps().isVal()){
								
							} else {
								index = i-1;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getCaps()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getCaps()!=null 
										&& firstRStyle.getRPr().getCaps().isVal()){
									chk1 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(r.getRPr().getCaps()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getCaps()!=null
										&& rStyle.getRPr().getCaps().isVal()){
									chk2 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								index = i-1;
								break;
							}
						}
						if(rPr.getSmallCaps()!=null && r.getRPr().getSmallCaps()!=null){
							if(rPr.getSmallCaps().isVal() && r.getRPr().getSmallCaps().isVal()){
								
							} else {
								index = i-1;
								break;
							}
						} else {
							boolean chk1 = false, chk2 = false;
							if(rPr.getSmallCaps()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null 
										&& firstRStyle.getRPr().getSmallCaps()!=null 
										&& firstRStyle.getRPr().getSmallCaps().isVal()){
									chk1 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(r.getRPr().getSmallCaps()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getSmallCaps()!=null
										&& rStyle.getRPr().getSmallCaps().isVal()){
									chk2 = true;
								} else {
									index = i-1;
									break;
								}
							}
							if(chk1&&chk2){
								
							} else {
								index = i-1;
								break;
							}
						}
						if(rPr.getSz()!=null && r.getRPr().getSz()!=null){
							if(rPr.getSz().getVal()!=null && r.getRPr().getSz().getVal()!=null
									&& rPr.getSz().getVal().intValue()==r.getRPr().getSz().getVal().intValue()){
							} else if(rPr.getSz().getVal()==null || r.getRPr().getSz().getVal()==null) {
								
							} else {
								index = i-1;
								break;
							}
						} else {
							BigInteger chk1 = null;
							BigInteger chk2 = null;
							if(rPr.getSz()==null){
								if(firstRStyle!=null && firstRStyle.getRPr()!=null
										&& firstRStyle.getRPr().getSz()!=null){
									chk1 = firstRStyle.getRPr().getSz().getVal();
								} else {
									index = i-1;
									break;
								}
							}
							if(r.getRPr()!=null && r.getRPr().getSz()==null){
								if(rStyle!=null && rStyle.getRPr()!=null
										&& rStyle.getRPr().getSz()!=null){
									chk2 = rStyle.getRPr().getSz().getVal();
								} else {
									index = i-1;
									break;
								}
							}
							if(chk1!=null && chk2!=null && chk1.intValue() == chk2.intValue()){
								
							} else {
								index = i-1;
								break;
							}
						}
					} else if(rPr==null && r.getRPr()==null){
						
					} else {
						index = i-1;
						break;
					}
				}
			}
		}
		if(index>=0){
			for(int i =0; i<= index; i++){
				Object o = pContents.get(i);
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(Context.getWmlObjectFactory().createRPr());
					r.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					r.getRPr().getRStyle().setVal(style.getStyleId());
				}
			}
		}
	}
	
	public static void replaceFirstSameGroupRStyle(P p, Style rStyle, Style style, String oldString, String addString, boolean test){
		boolean isAccordingStyle = false;
		Style xStyle = new Style();
		if(p!=null){
			R newR = null;
			if(addString!=null){
				newR = Context.getWmlObjectFactory().createR();
				Text text = Context.getWmlObjectFactory().createText();
				text.setValue(addString);
				newR.getContent().add(text);
				newR.setRPr(Context.getWmlObjectFactory().createRPr());
				newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
				newR.getRPr().getRStyle().setVal(style.getStyleId());
				newR.getContent().add(Context.getWmlObjectFactory().createRLastRenderedPageBreak());
			}
			
			int i=0;
			while(i<p.getContent().size()){
				Object o = p.getContent().get(i);
				if(o instanceof R){
					R r = (R)o;
					if(r.getRPr()!=null && rStyle!=null){
						xStyle.setRPr(r.getRPr());
						if(r.getRPr().getRStyle()!=null && rStyle.getRPr()!=null && rStyle.getRPr().getRStyle()!=null && r.getRPr().getRStyle().getVal().equals(rStyle.getRPr().getRStyle().getVal())){
							r.getRPr().getRStyle().setVal(style.getStyleId());
						} else {
							if(xStyle!=null){
								if(xStyle.getRPr().getB()!=null && rStyle.getRPr().getB()!=null){
									if(xStyle.getRPr().getB().isVal() && rStyle.getRPr().getB().isVal()){
										
									} else {
										break;
									}
								}
								if(xStyle.getRPr().getI()!=null && rStyle.getRPr().getI()!=null){
									if(xStyle.getRPr().getI().isVal() && rStyle.getRPr().getI().isVal()){
										
									} else {
										break;
									}
								}
								r.setRPr(Context.getWmlObjectFactory().createRPr());
								r.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
								r.getRPr().getRStyle().setVal(style.getStyleId());
								r.getContent().add(Context.getWmlObjectFactory().createRLastRenderedPageBreak());
							}
						}
					} else {
						if(rStyle == null && r.getRPr()==null){
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
							r.getRPr().getRStyle().setVal(style.getStyleId());
							r.getContent().add(Context.getWmlObjectFactory().createRLastRenderedPageBreak());
							
						} else {
							break;
						}
					}
				}
				i++;
			}
			if(oldString!=null && addString!=null){
				Object o = p.getContent().get(i-1);
				if(o instanceof R){
					R r = (R)o;
					String rText = getRText(r);
					int pos = rText.lastIndexOf(oldString);
					if(pos<0){
						p.getContent().add(i, newR);
					} else {
						rText = rText.substring(0, pos)+addString;
						r.getContent().clear();
						Text t = Context.getWmlObjectFactory().createText();
						t.setValue(rText);
						r.getContent().add(t);
					}
				}
			}
		}
	}
	
	public static boolean isPStyleFontProperitesAccepted(P p, Style chkStyle, StyleDefinitionsPart stylePart){
		boolean isAccepted = false;
		
		if(p!=null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
			Style pStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
			if(chkStyle.getRPr()!=null){
				if(pStyle!=null && pStyle.getRPr()!=null){
					if(pStyle.getRPr().getB()!=null && chkStyle.getRPr().getB()!=null){
						if(pStyle.getRPr().getB().isVal() && chkStyle.getRPr().getB().isVal()){
							isAccepted = true;
						}
					} else if(chkStyle.getRPr().getB()!=null) {
						if(!chkStyle.getRPr().getB().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					}
					if(pStyle.getRPr().getI()!=null && chkStyle.getRPr().getI()!=null){
						if(pStyle.getRPr().getI().isVal() && chkStyle.getRPr().getI().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					} else if(chkStyle.getRPr().getI()!=null) {
						if(!chkStyle.getRPr().getI().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					}
					if(pStyle.getRPr().getCaps()!=null && chkStyle.getRPr().getCaps()!=null){
						if(pStyle.getRPr().getCaps().isVal() && chkStyle.getRPr().getCaps().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					} else if(chkStyle.getRPr().getCaps()!=null) {
						if(!chkStyle.getRPr().getCaps().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					}
					if(pStyle.getRPr().getSmallCaps()!=null && chkStyle.getRPr().getSmallCaps()!=null){
						if(pStyle.getRPr().getSmallCaps().isVal() && chkStyle.getRPr().getSmallCaps().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
						}
					} else if(chkStyle.getRPr().getSmallCaps()!=null) {
						if(!chkStyle.getRPr().getSmallCaps().isVal()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					}
					if(pStyle.getRPr().getSz()!=null && chkStyle.getRPr().getSz()!=null){
						if(pStyle.getRPr().getSz().getVal()!=null && chkStyle.getRPr().getSz().getVal()!=null){
							if(pStyle.getRPr().getSz().getVal().intValue() == chkStyle.getRPr().getSz().getVal().intValue()){
								isAccepted = true;
							} else {
								isAccepted = false;
								return isAccepted;
							}
						} else {
							isAccepted = false;
							return isAccepted;
						}
					} else if(pStyle.getRPr().getSzCs()!=null && pStyle.getRPr().getSzCs().getVal()!=null && chkStyle.getRPr().getSz()!=null && chkStyle.getRPr().getSz().getVal()!=null) {
						if(pStyle.getRPr().getSzCs().getVal().intValue() == chkStyle.getRPr().getSz().getVal().intValue()){
							isAccepted = true;
						} else {
							isAccepted = false;
							return isAccepted;
						}
					}
				} else {
					isAccepted = false;
				}
			} else {
				
			}
		}
		
		return isAccepted;
	}
	
	public static void fixFirstLineEmptySpaces(P p, boolean isRemove, Style style){
		if(p!=null){
			if(p.getPPr()!=null && p.getPPr().getRPr()!=null){
				if(p.getPPr().getInd()!=null && p.getPPr().getInd().getFirstLine()!=null){
					p.getPPr().getInd().setFirstLine(null);
				}
			} else if(p.getPPr()!=null && p.getPPr().getPStyle()!=null) {
				if(isRemove){
					p.getPPr().setPStyle(null);
				} else {
					if(style!=null && style.getStyleId()!=null){
						p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						p.getPPr().getPStyle().setVal(style.getStyleId());
					}
				}
			}
		}
	}
	
	public static void removeSectPr(P p){
		if(p!=null){
			if(p.getPPr()!=null && p.getPPr().getSectPr()!=null){
				p.getPPr().setSectPr(null);
			}
		}
	}
	
	public static void replaceSectPr(P p, Style style){
		if(p!=null){
			if(p.getPPr()!=null && p.getPPr().getSectPr()!=null
					&& style.getPPr()!=null && style.getPPr().getSectPr()!=null){
				SectPr sectPr = p.getPPr().getSectPr();
				sectPr.setPgSz(style.getPPr().getSectPr().getPgSz());
				sectPr.setPgMar(style.getPPr().getSectPr().getPgMar());
//				sectPr.setType(style.getPPr().getSectPr().getType());
			}
		}
	}
	
	public static void setAlignmentCenter(P p){
		if(p!=null){
			if(p.getPPr()!=null){
				
			} else {
				p.setPPr(Context.getWmlObjectFactory().createPPr());
			}
			p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
			p.getPPr().getJc().setVal(JcEnumeration.CENTER);
		}
	}
	
	public static void removeInd(P p){
		if(p!=null){
			if(p.getPPr()!=null){
				
			} else {
				p.setPPr(Context.getWmlObjectFactory().createPPr());
			}
			p.getPPr().setInd(null);
		}
	}
	
	public static void fixFramePr(P p){
		if(p!=null && p.getPPr()!=null && p.getPPr().getFramePr()!=null){
			p.getPPr().getFramePr().setHRule(STHeightRule.AUTO);
			p.getPPr().getFramePr().setXAlign(STXAlign.CENTER);
//			p.getPPr().getFramePr().setYAlign(STYAlign.CENTER);
//			p.getPPr().getFramePr().setX(null);
			p.getPPr().getFramePr().setY(null);
//			p.getPPr().getFramePr().setVAnchor(null);
//			p.getPPr().getFramePr().setHAnchor(null);
//			p.getPPr().getFramePr().setH(null);
//			p.getPPr().getFramePr().setW(null);
			p.getPPr().getFramePr().setWrap(STWrap.NOT_BESIDE);
			p.getPPr().getFramePr().setHSpace(null);
			p.getPPr().getFramePr().setVSpace(null);
			p.getPPr().getFramePr().setAnchorLock(null);
		}
	}
	
	public static void convertUppcaseToNormalWord(P p){
		if(p!=null){
			String pText = getPText(p);
			if(pText.toUpperCase().equals(pText)){
				pText = pText.substring(0, 1)+pText.substring(1, pText.length()).toLowerCase();
				Text text = Context.getWmlObjectFactory().createText();
				text.setValue(pText);
				R r = Context.getWmlObjectFactory().createR();
				r.getContent().add(text);
				p.getContent().clear();
				p.getContent().add(r);
			}
		}
	}
	
	public static Style create4ChkingStyle(Style chkingStyle){
		chkingStyle = new Style();
		chkingStyle.setRPr(new RPr());
		chkingStyle.getRPr().setB(new BooleanDefaultTrue());
		chkingStyle.getRPr().getB().setVal(false);
		chkingStyle.getRPr().setI(new BooleanDefaultTrue());
		chkingStyle.getRPr().getI().setVal(false);
		chkingStyle.getRPr().setCaps(new BooleanDefaultTrue());
		chkingStyle.getRPr().getCaps().setVal(false);
		chkingStyle.getRPr().setSmallCaps(new BooleanDefaultTrue());
		chkingStyle.getRPr().getSmallCaps().setVal(false);
		return chkingStyle;
	}
	
	public static boolean isIncludeSectPr(P p){
		boolean result = false;
		if(p.getPPr()!=null && p.getPPr().getSectPr()!=null){
			result = true;
		} else {
			
		}
		return result;
	}
	
	public static void changeSectPr(P p, int cols, Style style){
		if(isIncludeSectPr(p) && style.getPPr()!=null && style.getPPr().getSectPr()!=null){
			SectPr sectPr = style.getPPr().getSectPr();
			sectPr.getCols().setNum(BigInteger.valueOf(cols));
			p.getPPr().setSectPr(sectPr);
		}
	}
	
	public static boolean isEmpty(P p){
		boolean isEmpty = false;
		String textStr = getPText(p);
		if(textStr!=null && textStr.length()>0){
			isEmpty = false;
		} else {
			isEmpty = true;
		}
		
		return isEmpty;
	}
	
	public static int removeNode(List<Object> contentList, int docContentIndex){
		
		if(contentList.get(docContentIndex)!=null){
			contentList.remove(docContentIndex);
			docContentIndex--;
		}
		
		return docContentIndex;
	}
	
	public static P removeFirstStr(P p, String str, String sep){
		String pText = getPText(p);
		if(pText!=null && pText.indexOf(str)>=0){
			pText = pText.replaceFirst(str, "");
			pText = pText.substring(pText.indexOf(sep)+1, pText.length());
			p.getContent().clear();
			R newR = Context.getWmlObjectFactory().createR();
			Text text = Context.getWmlObjectFactory().createText();
			text.setValue(pText);
			newR.getContent().add(text);
			p.getContent().add(newR);
		}
		
		return p;
	}
	
	public static P setPStyle(P p, String styleId){
		String pText = getPText(p);
		if(p!=null && pText!=null && pText.length()>0 && styleId!=null){
			PPr pPr = p.getPPr();
			if(pPr!=null){
				
			} else {
				pPr = Context.getWmlObjectFactory().createPPr();
			}
			pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
			pPr.getPStyle().setVal(styleId);
		}
		
		return p;
	}
	
	public static boolean isBibliography(SdtBlock sdtBlock){
		boolean isBibliography = false;
		if(sdtBlock.getSdtPr()!=null){
			if(sdtBlock.getSdtPr().getRPrOrAliasOrLock()!=null){
				for(Object x : sdtBlock.getSdtPr().getRPrOrAliasOrLock()){
					if(x instanceof JAXBElement<?> && ((JAXBElement<?>) x).getValue() instanceof Bibliography){
						isBibliography = true;
					}
				}
			}
		}
		return isBibliography;
	}
	
	public static boolean isTwoColumnTableAsRefFormat(Tbl tbl){
		boolean result = false;
		
		if(tbl.getTblGrid()!=null){
			if(tbl.getTblGrid().getGridCol()!=null && tbl.getTblGrid().getGridCol().size()==2){
				if(tbl.getTblGrid().getGridCol().get(0).getW().intValue()<800){
					result = true;
				}
			}
		}
				
		return result;
	}
	
	public static List<P> getGeneratedReferences(SdtBlock sdtBlock, String headStyleId, String styleId){
		List<P> result = new ArrayList<P>();
		try {
			boolean isBibliography = false;
			
			ContentAccessor sdtContent = sdtBlock.getSdtContent();
			if(sdtContent!=null){
				List<Object> sdtContentList = sdtContent.getContent();
				for(Object sdtContentO : sdtContentList){
					if(sdtContentO instanceof P){
						// reference head
						OOXMLConvertTool.removeAllProperties((P)sdtContentO);
						((P)sdtContentO).setPPr(Context.getWmlObjectFactory().createPPr());
						((P)sdtContentO).getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						((P)sdtContentO).getPPr().getPStyle().setVal(headStyleId);
						result.add((P)sdtContentO);
					} else if(sdtContentO instanceof SdtBlock) {
						SdtBlock sdtBlock2 = (SdtBlock) sdtContentO;
						isBibliography = isBibliography(sdtBlock2);
						if(isBibliography){
							System.out.println("yes yes");
						}
						ContentAccessor sdtContent2 = sdtBlock2.getSdtContent();
						if(sdtContent2!=null){
							List<Object> sdtContentList2 = sdtContent2.getContent();
							for(Object x: sdtContentList2){
								if(x instanceof SdtBlock){
									//
								} else if(x instanceof P){
									//
								} else if(x instanceof JAXBElement<?> && ((JAXBElement<?>)x).getValue() instanceof Tbl){
									Tbl tbl = (Tbl)((JAXBElement<?>)x).getValue();
									if(isTwoColumnTableAsRefFormat(tbl)){
										List<Object> tblContent = tbl.getContent();
										for(Object tblO : tblContent){
											if(tblO instanceof Tr){
												Tr tr = (Tr)tblO;
												List<Object> trContent = tr.getContent();
												int tcIndex = 0;
												for(Object trO : trContent){
													if(trO instanceof JAXBElement<?> && ((JAXBElement<?>)trO).getValue() instanceof Tc){
														if(tcIndex++>0){
															tcIndex = 0;
															Tc tc = (Tc)((JAXBElement<?>)trO).getValue();
															List<Object> tcContent = tc.getContent();
															for(Object tcO : tcContent){
																if(tcO instanceof P){
																	P p = (P)tcO;
																	OOXMLConvertTool.removeAllProperties(p);
																	p.setPPr(Context.getWmlObjectFactory().createPPr());
																	p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
																	p.getPPr().getPStyle().setVal(styleId);
																	result.add(p);
																}
															}
														}
													}
												}
											}
										}
									}
								} else {
									System.out.println(x.getClass().getName());
								}
							}
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return result;
	}
	
	public static int adaptGeneratedReferences(SdtBlock sdtBlock, String headStyleId, String styleId, int contentIndex, List<Object> contentList){
		int result = 0;
		try {
			List<P> refencesList = getGeneratedReferences(sdtBlock, headStyleId, styleId);
			
			if(contentList!=null && contentList.get(contentIndex)!=null){
				contentList.remove(contentIndex);
				for(int index=0; index<refencesList.size(); index++){
					contentList.add(contentIndex+index, refencesList.get(index));
				}
				result = refencesList.size();
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return result;
	}
}
