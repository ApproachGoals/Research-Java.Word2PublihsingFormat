package lib.ooxml.identifier;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.BooleanDefaultFalse;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTSimpleField;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.P;
import org.docx4j.wml.PPrBase.NumPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;
import org.docx4j.wml.DocDefaults.RPrDefault;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.Num;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import db.AcademicStyleMap;
import lib.ooxml.tool.OOXMLConvertTool;
import tools.LogUtils;

public class OOXMLIdentifier {

	private boolean isTitleIdentified;
	
	private boolean isSubTitleIdentified;
	
	private boolean isAbstractIdentified;
	
	private boolean isKeywordIdentified;
	
	private boolean isAcknowledgeIdentified;
	
	private boolean isReferencesHeaderIdentified;
	
	private boolean isAppendixHeaderIdentified;
	
	private RPrDefault rPrDefault;
	
	private String sourceFileFormat;
	
	public RPrDefault getrPrDefault() {
		return rPrDefault;
	}

	public void setrPrDefault(RPrDefault rPrDefault) {
		this.rPrDefault = rPrDefault;
	}

	public boolean isTitleIdentified() {
		return isTitleIdentified;
	}

	public void setTitleIdentified(boolean isTitleIdentified) {
		this.isTitleIdentified = isTitleIdentified;
	}

	public boolean isSubTitleIdentified() {
		return isSubTitleIdentified;
	}

	public void setSubTitleIdentified(boolean isSubTitleIdentified) {
		this.isSubTitleIdentified = isSubTitleIdentified;
	}

	public boolean isAbstractIdentified() {
		return isAbstractIdentified;
	}

	public void setAbstractIdentified(boolean isAbstractIdentified) {
		this.isAbstractIdentified = isAbstractIdentified;
	}

	public boolean isKeywordIdentified() {
		return isKeywordIdentified;
	}

	public void setKeywordIdentified(boolean isKeywordIdentified) {
		this.isKeywordIdentified = isKeywordIdentified;
	}

	public boolean isAcknowledgeIdentified() {
		return isAcknowledgeIdentified;
	}

	public void setAcknowledgeIdentified(boolean isAcknowledgeIdentified) {
		this.isAcknowledgeIdentified = isAcknowledgeIdentified;
	}

	public boolean isReferencesHeaderIdentified() {
		return isReferencesHeaderIdentified;
	}

	public void setReferencesHeaderIdentified(boolean isReferencesHeaderIdentified) {
		this.isReferencesHeaderIdentified = isReferencesHeaderIdentified;
	}

	public String getSourceFileFormat() {
		return sourceFileFormat;
	}

	public void setSourceFileFormat(String sourceFileFormat) {
		this.sourceFileFormat = sourceFileFormat;
	}

	public boolean isAppendixHeaderIdentified() {
		return isAppendixHeaderIdentified;
	}

	public void setAppendixHeaderIdentified(boolean isAppendixHeaderIdentified) {
		this.isAppendixHeaderIdentified = isAppendixHeaderIdentified;
	}

	public OOXMLIdentifier() {
		// TODO Auto-generated constructor stub
		sourceFileFormat = "";
	}
	
	public void init(StyleDefinitionsPart stylePart){
		try {
			List<Object> jaxbNodes = stylePart.getJAXBNodesViaXPath("w:docDefaults", false);
			if(jaxbNodes!=null){
				for (int i = 0; i < jaxbNodes.size(); i++) {
					if(jaxbNodes.get(i) instanceof DocDefaults){
						this.setrPrDefault(((DocDefaults) jaxbNodes.get(i)).getRPrDefault());
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
	}
	
	private boolean isPEmpty(P p){
		boolean result = false;
		
		if(getPText(p).length()==0){
			result = true;
		}
		
		return result;
	}
	
	private String getPText(P p){
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
						
					}
				}
			}
		}
		return result;
	}
	
	private String getRText(R r){
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
	
	private boolean chkRStyleInR(R r, Style chkStyle, StyleDefinitionsPart stylePart){
		boolean isAccepted = false;
		RPr rPr = null;
		if(r!=null && r.getRPr()!=null){
			rPr = r.getRPr();
		}
		String rText = getRText(r);
		if(rText!=null && rText.length()<=0){
			return true;
		}
		if(rPr != null && rPr.getRStyle()!=null && chkStyle!=null && chkStyle.getRPr()!=null){
			Style style = stylePart.getStyleById(rPr.getRStyle().getVal());
			if(chkStyle.getRPr().getB()!=null && style.getRPr().getB()!=null){
				if(chkStyle.getRPr().getB().isVal() && style.getRPr().getB().isVal()){
					isAccepted = true;
				} else {
					isAccepted = false;
					return isAccepted;
				}
			} else if(chkStyle.getRPr().getB()!=null && !chkStyle.getRPr().getB().isVal()){
				isAccepted = true;
			} else {
				
			}
			if(chkStyle.getRPr().getI()!=null && style.getRPr().getI()!=null){
				if(chkStyle.getRPr().getI().isVal() && style.getRPr().getI().isVal()){
					isAccepted = true;
				} else {
					isAccepted = false;
					return isAccepted;
				}
			} else if(chkStyle.getRPr().getI()!=null && !chkStyle.getRPr().getI().isVal()){
				isAccepted = true;
			} else {
				
			}
			if(chkStyle.getRPr().getSz()!=null && chkStyle.getRPr().getSz().getVal()!=null){
				if(style.getRPr().getSz()!=null && chkStyle.getRPr().getSz().getVal().compareTo(style.getRPr().getSz().getVal())==0){
					isAccepted = true;
				} else if(style.getRPr().getSzCs()!=null && chkStyle.getRPr().getSz().getVal().compareTo(style.getRPr().getSzCs().getVal())==0){
					isAccepted = true;
				} else {
					isAccepted = false;
					return isAccepted;
				}
			} else if(chkStyle.getRPr().getSz()==null && style.getRPr().getSz()==null) {
				isAccepted = true;
			} else {
				
			}
			if(chkStyle.getRPr().getRFonts()!=null && chkStyle.getRPr().getRFonts().getAscii()!=null){
				if(chkStyle.getRPr().getRFonts().getAscii().equals(style.getRPr().getRFonts().getAscii())){
					isAccepted = true;
				} else {
					isAccepted = false;
					return isAccepted;
				}
			}
			if(chkStyle.getRPr().getCaps()!=null && style.getRPr().getCaps()!=null){
				if(chkStyle.getRPr().getCaps().isVal() && style.getRPr().getCaps().isVal()){
					isAccepted = true;
				} else {
					isAccepted = false;
					return isAccepted;
				}
			} else if(chkStyle.getRPr().getCaps()!=null && !chkStyle.getRPr().getCaps().isVal()) {
				isAccepted = true;
			} else {
				
			}
			if(chkStyle.getRPr().getSmallCaps()!=null && style.getRPr().getSmallCaps()!=null){
				if(chkStyle.getRPr().getSmallCaps().isVal() && style.getRPr().getSmallCaps().isVal()){
					isAccepted = true;
				} else {
					isAccepted = false;
					return isAccepted;
				}
			} else if(chkStyle.getRPr().getSmallCaps()!=null && !chkStyle.getRPr().getSmallCaps().isVal()) {
				isAccepted = true;
			} else {
				
			}
		} else if(chkStyle!=null){
			if(chkStyle.getRPr()!=null){
				if(chkStyle.getRPr().getB()!=null && !chkStyle.getRPr().getB().isVal()
						|| chkStyle.getRPr().getI()!=null && !chkStyle.getRPr().getI().isVal()
						|| chkStyle.getRPr().getCaps()!=null && !chkStyle.getRPr().getCaps().isVal()
						|| chkStyle.getRPr().getSmallCaps()!=null && !chkStyle.getRPr().getSmallCaps().isVal()){
					isAccepted = true;
				} else {
					isAccepted = false;
				}
			} else {
				isAccepted = false;
			}
		} else {
			isAccepted = true;
		}
		
		return isAccepted;
	}
	
	private boolean chkStyleInR(P p, Style chkStyle, StyleDefinitionsPart stylePart){
		boolean isAccepted = false;
		
		String pText = getPText(p);
		if(pText!=null && pText.length()<=0){
			return true;
		}
		
		if(p!=null && p.getContent()!=null && p.getContent().size()>0 && chkStyle!=null){
			for (int i = 0; i < p.getContent().size(); i++) {
				Object o = p.getContent().get(i);
				if(o instanceof R){
					R r = (R)o;
					if(r.getRPr()!=null && chkStyle.getRPr()!=null){
						RPr rPr = r.getRPr();
						RPr chkRPr = chkStyle.getRPr();
						if(chkRPr.getB()!=null){
							if(rPr.getB()!=null){
								if(chkRPr.getB().isVal() == rPr.getB().isVal()){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getRStyle()!=null){
								Style docStyle = stylePart.getStyleById(rPr.getRStyle().getVal());
								ArrayList<String> appearedStyleList = new ArrayList<String>();
								appearedStyleList.add(rPr.getRStyle().getVal());
								isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								break;
							} else {
								String rText = getRText(r);
								if(rText!=null && rText.length()>0){
//									isAccepted = chkRStyleInR(r, chkStyle, stylePart);
									if(!isAccepted && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
										Style docStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(p.getPPr().getPStyle().getVal());
										isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
									} else if(rPr.getB()==null && !chkRPr.getB().isVal()) {
										isAccepted = true;
									}
									break;
								}
							}
						} else if(rPr.getB()!=null){
							isAccepted = false;
							break;
						}
						if(chkRPr.getI()!=null){
							if(rPr.getI()!=null){
								if(chkRPr.getI().isVal() == rPr.getI().isVal()){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getRStyle()!=null){
								Style docStyle = stylePart.getStyleById(rPr.getRStyle().getVal());
								ArrayList<String> appearedStyleList = new ArrayList<String>();
								appearedStyleList.add(rPr.getRStyle().getVal());
								isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								break;
							} else {
//								isAccepted = chkRStyleInR(r, chkStyle, stylePart);
								if(rPr.getI()==null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
									Style docStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
									ArrayList<String> appearedStyleList = new ArrayList<String>();
									appearedStyleList.add(p.getPPr().getPStyle().getVal());
									isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								} else if(rPr.getI()==null && !chkRPr.getI().isVal()) {
									isAccepted = true;
								}
								break;
							}
						} else if(rPr.getI()!=null) {
							isAccepted = false;
							break;
						}
						if(chkRPr.getRFonts()!=null && chkRPr.getRFonts().getAscii()!=null){
							if(rPr.getRFonts()!=null && rPr.getRFonts().getAscii()!=null){
								if(chkRPr.getRFonts().getAscii().equals(rPr.getRFonts().getAscii())){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getRStyle()!=null){
								Style docStyle = stylePart.getStyleById(rPr.getRStyle().getVal());
								ArrayList<String> appearedStyleList = new ArrayList<String>();
								appearedStyleList.add(rPr.getRStyle().getVal());
								isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								break;
							} else {
//								isAccepted = chkRStyleInR(r, chkStyle, stylePart);
								if(rPr.getRFonts()==null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
									Style docStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
									ArrayList<String> appearedStyleList = new ArrayList<String>();
									appearedStyleList.add(p.getPPr().getPStyle().getVal());
									isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								}
								break;
							}
						} else if(rPr.getRFonts()!=null){
							isAccepted = false;
							break;
						}
						if(chkRPr.getSz()!=null) {
							if(rPr.getSz()!=null && chkRPr.getSz().getVal()!=null && rPr.getSz().getVal()!=null){
								if(chkRPr.getSz().getVal().intValue() == rPr.getSz().getVal().intValue()){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getSzCs()!=null && chkRPr.getSz().getVal()!=null && rPr.getSzCs().getVal()!=null){
								if(chkRPr.getSz().getVal().intValue() == rPr.getSzCs().getVal().intValue()){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getRStyle()!=null){
								Style docStyle = stylePart.getStyleById(rPr.getRStyle().getVal());
								ArrayList<String> appearedStyleList = new ArrayList<String>();
								appearedStyleList.add(rPr.getRStyle().getVal());
								isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								break;
							} else {
								String rText = getRText(r);
								if(rText!=null && rText.length()>0){
//									isAccepted = chkRStyleInR(r, chkStyle, stylePart);
									if(rPr.getSz()==null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
										Style docStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(p.getPPr().getPStyle().getVal());
										isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
									}
									break;
								}
							}
//						} else if(rPr.getSz()!=null) {
//							isAccepted = false;
//							break;
						}
						if(chkRPr.getSmallCaps()!=null){
							if(rPr.getSmallCaps()!=null){
								if(chkRPr.getSmallCaps().isVal() && rPr.getSmallCaps().isVal()){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getRStyle()!=null){
								Style docStyle = stylePart.getStyleById(rPr.getRStyle().getVal());
								ArrayList<String> appearedStyleList = new ArrayList<String>();
								appearedStyleList.add(rPr.getRStyle().getVal());
								isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								break;
							} else {
								String rText = getRText(r);
								if(rText!=null && rText.length()>0){
//									isAccepted = chkRStyleInR(r, chkStyle, stylePart);
									if(rPr.getSmallCaps()==null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
										Style docStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(p.getPPr().getPStyle().getVal());
										isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
									} else if(rPr.getSmallCaps()==null && !chkRPr.getSmallCaps().isVal()) {
										isAccepted = true;
									}
									break;
								}
							}
						}
						if(chkRPr.getCaps()!=null){
							if(rPr.getCaps()!=null){
								if(chkRPr.getCaps().isVal() && rPr.getCaps().isVal()){
									isAccepted = true;
									break;
								} else {
									isAccepted = false;
									break;
								}
							} else if(rPr.getRStyle()!=null){
								Style docStyle = stylePart.getStyleById(rPr.getRStyle().getVal());
								ArrayList<String> appearedStyleList = new ArrayList<String>();
								appearedStyleList.add(rPr.getRStyle().getVal());
								isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
								break;
							} else {
								String rText = getRText(r);
								if(rText!=null && rText.length()>0){
//									isAccepted = chkRStyleInR(r, chkStyle, stylePart);
									if(rPr.getCaps()==null && p.getPPr()!=null && p.getPPr().getPStyle()!=null){
										Style docStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(p.getPPr().getPStyle().getVal());
										isAccepted = rekursiveChkStyle(stylePart, chkStyle, docStyle, appearedStyleList, true);
									} else if(rPr.getCaps()==null && !chkRPr.getCaps().isVal()) {
										isAccepted = true;
									}
									break;
								}
							}
						} else if(rPr.getCaps()!=null){
							isAccepted = false;
							break;
						}
//						if(!isAccepted){
//							if(chkRPr.getB()==null && rPr.getB()==null
//									&& chkRPr.getI()==null && rPr.getI()==null
//									&& chkRPr.getSmallCaps()==null && rPr.getSmallCaps()==null
//									&& chkRPr.getCaps()==null && rPr.getCaps()==null
//									&& chkRPr.getSz()==null && rPr.getSz()==null
//									){
//								isAccepted = chkRStyleInR(r, chkStyle, stylePart);
//								break;
//							}
//						}
					} else {
						if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
							isAccepted = OOXMLConvertTool.isPStyleFontProperitesAccepted(p, chkStyle, stylePart);
							break;
						}
					}
				}
			}
		} else if(chkStyle!=null){
			isAccepted = false;
		} else {
			isAccepted = true;
		}
		
		return isAccepted;
	}
	
	private boolean chkStyleSingleProperty(Style chkingStyle, Style style){
		boolean result = false;
		
		if(chkingStyle!=null && style!=null){
			if(chkingStyle.getRPr()!=null && style.getRPr()!=null){
				if(style.getRPr().getB()!=null && chkingStyle.getRPr().getB()!=null){
					if(chkingStyle.getRPr().getB().isVal() 
							&& style.getRPr().getB().isVal()){
						result = true;
						return result;
					} else {
						result = false;
						return result;
					}
				} else if(style.getRPr().getB()==null && chkingStyle.getRPr().getB()!=null && !chkingStyle.getRPr().getB().isVal()) {
					result = true;
				} else {
					result = false;//ignore
				}
				if(style.getRPr().getI()!=null && chkingStyle.getRPr().getI()!=null){
					if(chkingStyle.getRPr().getI().isVal() 
							&& style.getRPr().getI().isVal()){
						result = true;
						return result;
					} else {
						result = false;
						return result;
					}
				} else if(style.getRPr().getI()==null && chkingStyle.getRPr().getI()!=null && !chkingStyle.getRPr().getI().isVal()) {
					result = true;
				} else {
					result = false;//ignore
				}
				if(style.getRPr().getSmallCaps()!=null && chkingStyle.getRPr().getSmallCaps()!=null){
					if(chkingStyle.getRPr().getSmallCaps().isVal() 
							&& style.getRPr().getSmallCaps().isVal()){
						result = true;
						return result;
					} else {
						result = false;
						return result;
					}
				} else if(style.getRPr().getSmallCaps()==null && chkingStyle.getRPr().getSmallCaps()!=null && !chkingStyle.getRPr().getSmallCaps().isVal()) {
					result = true;
				} else {
					result = false;//ignore
				}
				if(style.getRPr().getCaps()!=null && chkingStyle.getRPr().getCaps()!=null){
					if(chkingStyle.getRPr().getCaps().isVal() 
							&& style.getRPr().getCaps().isVal()){
						result = true;
						return result;
					} else {
						result = false;
						return result;
					}
				} else if(style.getRPr().getCaps()==null && chkingStyle.getRPr().getCaps()!=null && !chkingStyle.getRPr().getCaps().isVal()) {
					result = true;
				} else {
					result = false;//ignore
				}
				
				if(style.getRPr().getSz()!=null && chkingStyle.getRPr().getSz()!=null){
					if(chkingStyle.getRPr().getSz().getVal()!=null
							&& style.getRPr().getSz().getVal()!=null
							&& chkingStyle.getRPr().getSz().getVal().intValue()==style.getRPr().getSz().getVal().intValue()){
						result = true;
						return result;
					} else {
						result = false;
						return result;
					}
//				} else if(chkingStyle.getRPr().getSz()==null){
//					result = true;
				} else {
//					result = false;//ignore
				}
				if(style.getRPr().getSz()!=null && chkingStyle.getRPr().getSzCs()!=null){
					if(chkingStyle.getRPr().getSzCs().getVal()!=null
							&& style.getRPr().getSz().getVal()!=null
							&& chkingStyle.getRPr().getSzCs().getVal().intValue()==style.getRPr().getSz().getVal().intValue()){
						result = true;
						return result;
					} else {
						
					}
				} else {
					
				}
//			} else if(chkingStyle.getRPr()==null && style.getRPr()==null) {
//				result = true;
			} else {
//				result = false;
			}
		} else if(chkingStyle==null && style == null){
			result = true;
		} 
		
		return result;
	}
	
	private boolean chkStyleMultiProperties(Style chkingStyle, Style style){
		boolean result = false;
		/*
		 * bold, italic, smallCaps, caps, sz
		 */
		if(chkingStyle!=null && style !=null){
			if(chkingStyle.getRPr()!=null && style.getRPr()!=null){
				if(style.getRPr().getB()!=null && chkingStyle.getRPr().getB()!=null){
					if(chkingStyle.getRPr().getB().isVal() 
							&& style.getRPr().getB().isVal()){
						result = true;
					} else {
						result = false;
						return result;
					}
//				} else if(style.getRPr().getB()==null && chkingStyle.getRPr().getB()==null) {
//					result = true;
				} else {
					result = false;//ignore
				}
				if(style.getRPr().getI()!=null && chkingStyle.getRPr().getI()!=null){
					if(chkingStyle.getRPr().getI().isVal() 
							&& style.getRPr().getI().isVal()){
						result = true;
					} else {
						result = false;
						return result;
					}
//				} else if(style.getRPr().getI()==null && chkingStyle.getRPr().getI()==null) {
//					result = true;
				} else {
					result = false;//ignore
				}
				if(style.getRPr().getSmallCaps()!=null && chkingStyle.getRPr().getSmallCaps()!=null){
					if(chkingStyle.getRPr().getSmallCaps().isVal() 
							&& style.getRPr().getSmallCaps().isVal()){
						result = true;
					} else {
						result = false;
						return result;
					}
//				} else if(style.getRPr().getSmallCaps()==null && chkingStyle.getRPr().getSmallCaps()==null) {
//					result = true;
				} else {
					result = false;//ignore
				}
				if(style.getRPr().getCaps()!=null && chkingStyle.getRPr().getCaps()!=null){
					if(chkingStyle.getRPr().getCaps().isVal() 
							&& style.getRPr().getCaps().isVal()){
						result = true;
					} else {
						result = false;
						return result;
					}
//				} else if(style.getRPr().getCaps()==null && chkingStyle.getRPr().getCaps()==null) {
//					result = true;
				} else {
					result = false;//ignore
				}
				
				if(style.getRPr().getSz()!=null && chkingStyle.getRPr().getSz()!=null){
					if(chkingStyle.getRPr().getSz().getVal()!=null
							&& style.getRPr().getSz().getVal()!=null
							&& chkingStyle.getRPr().getSz().getVal().intValue()==style.getRPr().getSz().getVal().intValue()){
						result = true;
					} else {
						result = false;
					}
//				} else if(style.getRPr().getSz()==null && chkingStyle.getRPr().getSz()==null){
//					result = true;
				} else {
//					result = false;//ignore
				}
				if(style.getRPr().getSz()!=null && chkingStyle.getRPr().getSzCs()!=null){
					if(chkingStyle.getRPr().getSzCs().getVal()!=null
							&& style.getRPr().getSz().getVal()!=null
							&& chkingStyle.getRPr().getSzCs().getVal().intValue()==style.getRPr().getSz().getVal().intValue()){
						result = true;
					} else {
						
					}
				} else {
					
				}
				
				if(style.getRPr().getB()==null && chkingStyle.getRPr().getB()==null
						&& style.getRPr().getI()==null && chkingStyle.getRPr().getI()==null
						&& style.getRPr().getSmallCaps()==null && chkingStyle.getRPr().getSmallCaps()==null
						&& style.getRPr().getCaps()==null && chkingStyle.getRPr().getCaps()==null
						&& style.getRPr().getSz()==null && chkingStyle.getRPr().getSz()==null
						&& style.getRPr().getRFonts()==null && chkingStyle.getRPr().getRFonts()==null) {
					result = true;
				}
			} else if(chkingStyle.getRPr()==null && style.getRPr()==null) {
				result = true;
			} else {
//				result = false;
			}
		}
		
		return result;
	}
	
	private boolean rekursiveChkStyle(StyleDefinitionsPart stylePart, Style chkingStyle, Style style, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		
		if(style==null){
			result = false;
		} else if(chkingStyle!=null){
			if(chkStyleSingleProperty(chkingStyle, style)){
				result = true;
			} else {
				if(style.getBasedOn()!=null){
					Style basedOnStyle = null;
					if(!isFromDoc){
						AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
						basedOnStyle = map.getStyleMap().get(style.getBasedOn().getVal());
					} else {
						basedOnStyle = stylePart.getStyleById(style.getBasedOn().getVal());
					}
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
						} else {
							appearedStyleList.add(basedOnStyle.getStyleId());
							result = rekursiveChkStyle(stylePart, chkingStyle, basedOnStyle, appearedStyleList, isFromDoc);
						}
					} else {
						result = false;
					}
				}
				if(!result && style.getLink()!=null) {
					Style linkStyle = null;
					if(!isFromDoc){
						AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
						linkStyle = map.getStyleMap().get(style.getLink().getVal());
					} else {
						linkStyle = stylePart.getStyleById(style.getLink().getVal());
					}
					if(linkStyle!=null){
						if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
							result = false;
						} else {
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkStyle(stylePart, chkingStyle, linkStyle, appearedStyleList, isFromDoc);
						}
					} else {
						result = false;
					}
				}
			}
		} else {
			result = true;
		}
		
		return result;
	}
	/*
	private boolean isFontSzSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean result = false;
		if(predefinedStyle!=null && predefinedStyle.getRPr()!=null 
				&& predefinedStyle.getRPr().getSz()!=null && predefinedStyle.getRPr().getSz().getVal()!=null){
			
		} else {
			return true;
		}
//		if(p.getContent()!=null && p.getContent().size()>0 && isDeepChk) {
//			Style chkStyle = new Style();
//			chkStyle.setRPr(new RPr());
//			chkStyle.getRPr().setSz(predefinedStyle.getRPr().getSz());
//			result = chkStyleInR(p, chkStyle, stylePart);
//			if(result) return result;
//			else return false;
//		}
		if(p!=null && p.getPPr()!=null) {
			if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSz()!=null 
					&& p.getPPr().getRPr().getSz().getVal()!=null){
				if(p.getPPr().getRPr().getSz().getVal().intValue()==predefinedStyle.getRPr().getSz().getVal().intValue()){
					result = true;
				}
			} else if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSzCs()!=null 
					&& p.getPPr().getRPr().getSzCs().getVal()!=null){
				if(p.getPPr().getRPr().getSzCs().getVal().intValue()==predefinedStyle.getRPr().getSz().getVal().intValue()){
					result = true;
				}
			} else if(p.getPPr().getPStyle()!=null) {
				Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
//				Style basedOnStyle = style.getBasedOn()!=null?stylePart.getStyleById(style.getBasedOn().getVal()):null;
//				Style linkStyle = style.getLink()!=null?stylePart.getStyleById(style.getLink().getVal()):null;
//				
//				if(style.getRPr()!=null && style.getRPr().getSz()!=null && style.getRPr().getSz().getVal()!=null){
//					if(style.getRPr().getSz().getVal().intValue()== predefinedStyle.getRPr().getSz().getVal().intValue()){
//						result = true;
//					}
//				} else if(style.getRPr()!=null && style.getRPr().getSzCs()!=null && style.getRPr().getSzCs().getVal()!=null){
//					if(style.getRPr().getSzCs().getVal().intValue()==predefinedStyle.getRPr().getSz().getVal().intValue()){
//						result = true;
//					}
//				}
				if(style!=null){
					Style chkingStyle = OOXMLConvertTool.create4ChkingStyle(new Style());
					chkingStyle.setRPr(new RPr());
					chkingStyle.getRPr().setSz(predefinedStyle.getRPr().getSz());
					ArrayList<String> appearedStyleList = new ArrayList<String>();
					appearedStyleList.add(style.getStyleId());
					result = rekursiveChkStyle(stylePart, chkingStyle, style, appearedStyleList, true);
				}
				
				
//				} else if(basedOnStyle!=null && basedOnStyle.getRPr()!=null && basedOnStyle.getRPr().getSz()!=null && basedOnStyle.getRPr().getSz().getVal()!=null){
//					if(basedOnStyle.getRPr().getSz().getVal().intValue()== predefinedStyle.getRPr().getSz().getVal().intValue()){
//						result = true;
//					} else if(basedOnStyle.getLink()!=null){
//						Style linkedStyle = stylePart.getStyleById(basedOnStyle.getLink().getVal());
//						if(linkedStyle.getRPr()!=null && linkedStyle.getRPr().getSz()!=null && linkedStyle.getRPr().getSz().getVal()!=null) {
//							if(linkedStyle.getRPr().getSz().getVal().intValue() == predefinedStyle.getRPr().getSz().getVal().intValue()){
//								result = true;
//							} else if(linkedStyle.getRPr().getSzCs()!=null && linkedStyle.getRPr().getSzCs().getVal().intValue() == predefinedStyle.getRPr().getSz().getVal().intValue()){
//								result = true;
//							}
//						}
//					}
//				} else if(linkStyle!=null && linkStyle.getRPr()!=null && linkStyle.getRPr().getSz()!=null && linkStyle.getRPr().getSz().getVal()!=null) {
//					Style linkBasedOnStyle = stylePart.getStyleById(linkStyle.getBasedOn().getVal());
//					if(linkStyle.getRPr().getSz().getVal().intValue() == predefinedStyle.getRPr().getSz().getVal().intValue()){
//						result = true;
//					} else if(linkStyle.getRPr().getSzCs() != null && linkStyle.getRPr().getSzCs().getVal().intValue() == predefinedStyle.getRPr().getSz().getVal().intValue()){
//						result = true;
//					} else if(linkBasedOnStyle!=null){
//						
//					}
//					
//				}
			}
		}
		if(!result){
			if(stylePart.getDefaultParagraphStyle()!=null 
					&& stylePart.getDefaultParagraphStyle().getPPr()!=null
					&& stylePart.getDefaultParagraphStyle().getPPr().getRPr()!=null){
				if(stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz().getVal()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSz().getVal().intValue()==predefinedStyle.getRPr().getSz().getVal().intValue()
						){
					result = true;
				} else if(stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSzCs()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSzCs().getVal()!=null
						&& stylePart.getDefaultParagraphStyle().getPPr().getRPr().getSzCs().getVal().intValue()==predefinedStyle.getRPr().getSz().getVal().intValue()
						){
					result = true;
				}
			}
		}
		
		return result;
	}*/
	private boolean isFontSzSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean result = false;
		
		if(p!=null){
			OOXMLStyleLvlIdentifier styleIdentifier = new OOXMLStyleLvlIdentifier(this.sourceFileFormat, this.getrPrDefault());
			OOXMLStylePLvlIdentifier pIdentifier = new OOXMLStylePLvlIdentifier(this.sourceFileFormat);
			OOXMLStyleRLvlIdentifier rIdentifier = new OOXMLStyleRLvlIdentifier(this.sourceFileFormat);
			OOXMLStyleFromNumberingIdentifier nIdentifier = new OOXMLStyleFromNumberingIdentifier(sourceFileFormat);
//			Style style = new Style();
//			style.setRPr(new RPr());
//			if(predefinedStyle.getRPr().getSz()!=null && predefinedStyle.getRPr().getSz().getVal()!=null){
//				style.getRPr().setSz(predefinedStyle.getRPr().getSz());
//			} else {
//				style.getRPr().setSz(null);
//			}
			
//			boolean isRLvlPropertiesAccepted = false;
			UserMessage msg = new UserMessage();
			msg = rIdentifier.isFontSzAccepted(p, stylePart, predefinedStyle, msg);
			if(msg.isFullySuitable()){
//				isRLvlPropertiesAccepted = true;
				result = true;
			} else {
				if(isDeepChk && msg.isChkresultSuitable()){
//					isRLvlPropertiesAccepted = true;
					result = true;
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
//				boolean isRLvlStyleAccepted = false;// = styleIdentifier.isBAccepted(p, stylePart, style, predefinedStyle, msg);
				msg = new UserMessage();
				msg = styleIdentifier.isRFontSzAccepted(p, stylePart, predefinedStyle, isDeepChk);
				if(msg.isFullySuitable()){
//					isRLvlStyleAccepted = true;
					result = true;
				} else {
					if(isDeepChk && msg.isChkresultSuitable()){
//						isRLvlStyleAccepted = true;
						result = true;
					}
				}
//				if(!isRLvlStyleAccepted){
//					
//				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
//				boolean isPLvlPropertiesAccepted = false; 
				msg = new UserMessage();
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				msg = pIdentifier.isFontSzAccepted(p, stylePart, predefinedStyle, msg, appearedStylelist, false);
				if(msg.isFullySuitable()){
//					isPLvlPropertiesAccepted = true;
					result = true;
				} else {
					if(isDeepChk && msg.isChkresultSuitable()){
//						isPLvlPropertiesAccepted = true;
						result = true;
					}
				}
//				if(!isPLvlPropertiesAccepted){
//					
//				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
//				boolean isPLvlStyleAccepted = false; 
				msg = new UserMessage();
				msg = styleIdentifier.isFontSzAccepted(p, stylePart, predefinedStyle, msg);
				if(msg.isFullySuitable()){
//					isPLvlStyleAccepted = true;
					result = true;
				} else {
					if(isDeepChk && msg.isChkresultSuitable()){
//						isPLvlStyleAccepted = true;
						result = true;
					}
				}
//				if(!isPLvlStyleAccepted){
//					
//				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				msg = new UserMessage();
				msg = nIdentifier.isFontSzAccepted(p, stylePart, predefinedStyle, msg);
				if(msg.isFullySuitable()){
					result = true;
				} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
					result = true;
				}
			}
		}
		
		return result;
	}
//	public boolean isFontItalicSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
//		boolean result = false;
//		if(isDeepChk){
//			Style style = new Style();
//			style.setRPr(new RPr());
//			if(predefinedStyle.getRPr().getI()!=null && predefinedStyle.getRPr().getI().isVal()){
//				style.getRPr().setI(new BooleanDefaultTrue());
//			} else {
//				BooleanDefaultTrue val = new BooleanDefaultTrue();
//				val.setVal(false);
//				style.getRPr().setI(val);
//			}
//			result = chkStyleInR(p, style, stylePart);
//			if(result) return result;
//			else return false;
//		}		
//		if(predefinedStyle!=null){
//			Style chkingStyle = OOXMLConvertTool.create4ChkingStyle(new Style());
//			chkingStyle.getRPr().getI().setVal(true);
//			ArrayList<String> appearedStyleList = new ArrayList<String>();
//			appearedStyleList.add(predefinedStyle.getStyleId());
//			boolean isShouldItalic = rekursiveChkStyle(stylePart, chkingStyle, predefinedStyle, appearedStyleList, false);
//			if(p!=null && p.getPPr()!=null){
//				Style style = p.getPPr().getPStyle()!=null?stylePart.getStyleById(p.getPPr().getPStyle().getVal()):null;
////				Style basedOnStyle = style.getBasedOn()!=null?stylePart.getStyleById(style.getBasedOn().getVal()):null;
////				Style linkStyle = style.getLink()!=null?stylePart.getStyleById(style.getLink().getVal()):null;
//				Style rStyle = OOXMLConvertTool.getFirstRStyle(p);
//				if(rStyle != null){
//					style = rStyle;
//				}
//				if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getI()!=null) {
//					if(p.getPPr().getRPr().getI().isVal() && isShouldItalic){
//						result = true;
//					}
//				} else if(style!=null){
//					appearedStyleList = new ArrayList<String>();
//					appearedStyleList.add(style.getStyleId());
//					result = rekursiveChkStyle(stylePart, chkingStyle, style, appearedStyleList, true) == isShouldItalic;
//				}
//			} else {
//				result = true;
//			}
//		} else {
//			result = true;
//		}
//		
//		return result;
//	}
	
	public boolean isFontBoldSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean result = false;
//		if(isDeepChk){
//			Style style = new Style();
//			style.setRPr(new RPr());
//			if(predefinedStyle.getRPr().getB()!=null && predefinedStyle.getRPr().getB().isVal()){
//				style.getRPr().setB(new BooleanDefaultTrue());
//			} else {
//				BooleanDefaultTrue val = new BooleanDefaultTrue();
//				val.setVal(false);
//				style.getRPr().setB(val);
//			}
//			result = chkStyleInR(p, style, stylePart);
//			if(result) return result;
//			else return false;
//		}		
//		if(predefinedStyle!=null){
//			Style chkingStyle = OOXMLConvertTool.create4ChkingStyle(new Style());
//			chkingStyle.getRPr().getB().setVal(true);
//			ArrayList<String> appearedStyleList = new ArrayList<String>();
//			appearedStyleList.add(predefinedStyle.getStyleId());
//			boolean isShouldBold = rekursiveChkStyle(stylePart, chkingStyle, predefinedStyle, appearedStyleList, false);
//			if(p!=null && p.getPPr()!=null){
//				Style style = p.getPPr().getPStyle()!=null?stylePart.getStyleById(p.getPPr().getPStyle().getVal()):null;
//				Style rStyle = OOXMLConvertTool.getFirstRStyle(p);
//				if(rStyle != null){
//					style = rStyle;
//				}
//				if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getB()!=null) {
//					if(p.getPPr().getRPr().getB().isVal() && isShouldBold){
//						result = true;
//					}
//				} else if(style!=null){
//					appearedStyleList = new ArrayList<String>();
//					appearedStyleList.add(style.getStyleId());
//					result = rekursiveChkStyle(stylePart, chkingStyle, style, appearedStyleList, true) == isShouldBold;
//				}
//			} else {
//				result = true;
//			}
//		} else {
//			result = true;
//		}
		
		if(p!=null){
			OOXMLStyleLvlIdentifier styleIdentifier = new OOXMLStyleLvlIdentifier(this.sourceFileFormat, this.getrPrDefault());
			OOXMLStylePLvlIdentifier pIdentifier = new OOXMLStylePLvlIdentifier(this.sourceFileFormat);
			OOXMLStyleRLvlIdentifier rIdentifier = new OOXMLStyleRLvlIdentifier(this.sourceFileFormat);
			
			boolean isRLvlPropertiesAccepted = false;
			UserMessage msg = new UserMessage();
			msg = rIdentifier.isBAccepted(p, stylePart, predefinedStyle, msg);
			if(msg.isFullySuitable()){
				isRLvlPropertiesAccepted = true;
				result = true;
			} else {
				if(isDeepChk && msg.isChkresultSuitable()){
					isRLvlPropertiesAccepted = true;
					result = true;
				} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
					isRLvlPropertiesAccepted = true;
					result = true;
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isRLvlStyleAccepted = false;// = styleIdentifier.isBAccepted(p, stylePart, style, predefinedStyle, msg);
				msg = new UserMessage();
				msg = styleIdentifier.isRBAccepted(p, stylePart, predefinedStyle, isDeepChk);
				if(!isRLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isRLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isRLvlStyleAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlPropertiesAccepted = false; 
				msg = new UserMessage();
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				msg = pIdentifier.isBAccepted(p, stylePart, predefinedStyle, msg, appearedStylelist, false);
				if(!isPLvlPropertiesAccepted){
					if(msg.isFullySuitable()){
						isPLvlPropertiesAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlPropertiesAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlStyleAccepted = false;
				msg = new UserMessage();
				msg = styleIdentifier.isBAccepted(p, stylePart, predefinedStyle, msg);
				if(!isPLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isPLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isStyleFromNumbering = false;
			}
		}
		
		return result;
	}
	public boolean isFontItalicSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean result = false;
		
		if(p!=null){
			OOXMLStyleLvlIdentifier styleIdentifier = new OOXMLStyleLvlIdentifier(this.sourceFileFormat, this.getrPrDefault());
			OOXMLStylePLvlIdentifier pIdentifier = new OOXMLStylePLvlIdentifier(this.sourceFileFormat);
			OOXMLStyleRLvlIdentifier rIdentifier = new OOXMLStyleRLvlIdentifier(this.sourceFileFormat);
			
			boolean isRLvlPropertiesAccepted = false;
			UserMessage msg = new UserMessage();
			msg = rIdentifier.isIAccepted(p, stylePart, predefinedStyle, msg);
			if(msg.isFullySuitable()){
				isRLvlPropertiesAccepted = true;
				result = true;
			} else {
				if(isDeepChk && msg.isChkresultSuitable()){
					isRLvlPropertiesAccepted = true;
					result = true;
				} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
					isRLvlPropertiesAccepted = true;
					result = true;
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isRLvlStyleAccepted = false;// = styleIdentifier.isBAccepted(p, stylePart, style, predefinedStyle, msg);
				msg = new UserMessage();
				msg = styleIdentifier.isRIAccepted(p, stylePart, predefinedStyle, isDeepChk);
				if(!isRLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isRLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isRLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlPropertiesAccepted = false; 
				msg = new UserMessage();
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				msg = pIdentifier.isIAccepted(p, stylePart, predefinedStyle, msg, appearedStylelist, false);
				if(!isPLvlPropertiesAccepted){
					if(msg.isFullySuitable()){
						isPLvlPropertiesAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlPropertiesAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlStyleAccepted = false; 
				msg = new UserMessage();
//				String pText = OOXMLConvertTool.getPText(p);
				msg = styleIdentifier.isIAccepted(p, stylePart, predefinedStyle, msg);
				if(!isPLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isPLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
		}
		
		return result;
	}
	public boolean isFontCapsSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean result = false;
		
		if(p!=null){
			OOXMLStyleLvlIdentifier styleIdentifier = new OOXMLStyleLvlIdentifier(this.sourceFileFormat, this.getrPrDefault());
			OOXMLStylePLvlIdentifier pIdentifier = new OOXMLStylePLvlIdentifier(this.sourceFileFormat);
			OOXMLStyleRLvlIdentifier rIdentifier = new OOXMLStyleRLvlIdentifier(this.sourceFileFormat);
			
			boolean isRLvlPropertiesAccepted = false;
			UserMessage msg = new UserMessage();
			msg = rIdentifier.isCapsAccepted(p, stylePart, predefinedStyle, msg);
			if(msg.isFullySuitable()){
				isRLvlPropertiesAccepted = true;
				result = true;
			} else {
				if(isDeepChk && msg.isChkresultSuitable()){
					isRLvlPropertiesAccepted = true;
					result = true;
				} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
					isRLvlPropertiesAccepted = true;
					result = true;
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isRLvlStyleAccepted = false;// = styleIdentifier.isBAccepted(p, stylePart, style, predefinedStyle, msg);
				msg = styleIdentifier.isRCapsAccepted(p, stylePart, predefinedStyle, isDeepChk);
				if(!isRLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isRLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isRLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlPropertiesAccepted = false; 
				msg = new UserMessage();
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				msg = pIdentifier.isCapsAccepted(p, stylePart, predefinedStyle, msg, appearedStylelist, false);
				if(!isPLvlPropertiesAccepted){
					if(msg.isFullySuitable()){
						isPLvlPropertiesAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlPropertiesAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlStyleAccepted = false; 
				msg = new UserMessage();
				msg = styleIdentifier.isCapsAccepted(p, stylePart, predefinedStyle, msg);
				if(!isPLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isPLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
		}
		
		return result;
	}
	
//	public boolean isFontCapsSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
//		boolean result = false;
//		String pText = OOXMLConvertTool.getPText(p);
//		if(pText.toUpperCase().equals(pText)){
//			result = true;
//			return result;
//		}
//		if(isDeepChk){
//			Style style = new Style();
//			style.setRPr(new RPr());
//			if(predefinedStyle.getRPr().getCaps()!=null && predefinedStyle.getRPr().getCaps().isVal()){
//				style.getRPr().setCaps(new BooleanDefaultTrue());
//			} else {
//				BooleanDefaultTrue val = new BooleanDefaultTrue();
//				val.setVal(false);
//				style.getRPr().setCaps(val);
//			}
//			result = chkStyleInR(p, style, stylePart);
//			if(result) return result;
//			else return false;
//		}		
//		if(predefinedStyle!=null){
//			Style chkingStyle = OOXMLConvertTool.create4ChkingStyle(new Style());
//			chkingStyle.getRPr().getCaps().setVal(true);
//			ArrayList<String> appearedStyleList = new ArrayList<String>();
//			appearedStyleList.add(predefinedStyle.getStyleId());
//			boolean isShouldCaps = rekursiveChkStyle(stylePart, chkingStyle, predefinedStyle, appearedStyleList, false);
//			if(p!=null && p.getPPr()!=null){
//				Style style = p.getPPr().getPStyle()!=null?stylePart.getStyleById(p.getPPr().getPStyle().getVal()):null;
//				Style rStyle = OOXMLConvertTool.getFirstRStyle(p);
//				if(rStyle != null){
//					style = rStyle;
//				}
//				if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getCaps()!=null) {
//					if(p.getPPr().getRPr().getCaps().isVal() && isShouldCaps){
//						result = true;
//					} else {
//						
//					}
//				} else if(style!=null){
//					appearedStyleList = new ArrayList<String>();
//					appearedStyleList.add(style.getStyleId());
//					result = rekursiveChkStyle(stylePart, chkingStyle, style, appearedStyleList, true) == isShouldCaps;
//				}
//			} else {
//				result = true;
//			}
//		} else {
//			result = true;
//		}
//		
//		return result;
//	}
	public boolean isFontSmallCapsSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
		boolean result = false;
		
		if(p!=null){
			OOXMLStyleLvlIdentifier styleIdentifier = new OOXMLStyleLvlIdentifier(this.sourceFileFormat, this.getrPrDefault());
			OOXMLStylePLvlIdentifier pIdentifier = new OOXMLStylePLvlIdentifier(this.sourceFileFormat);
			OOXMLStyleRLvlIdentifier rIdentifier = new OOXMLStyleRLvlIdentifier(this.sourceFileFormat);
			
			boolean isRLvlPropertiesAccepted = false;
			UserMessage msg = new UserMessage();
			msg = rIdentifier.isSmallCapsAccepted(p, stylePart, predefinedStyle, msg);
			if(msg.isFullySuitable()){
				isRLvlPropertiesAccepted = true;
				result = true;
			} else {
				if(isDeepChk && msg.isChkresultSuitable()){
					isRLvlPropertiesAccepted = true;
					result = true;
				} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
					isRLvlPropertiesAccepted = true;
					result = true;
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isRLvlStyleAccepted = false; 
				msg = new UserMessage();
				msg = styleIdentifier.isRSmallCapsAccepted(p, stylePart, predefinedStyle, isDeepChk);
				if(!isRLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isRLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isRLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlPropertiesAccepted = false; 
				msg = new UserMessage();
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				msg = pIdentifier.isSmallCapsAccepted(p, stylePart, predefinedStyle, msg, appearedStylelist, false);
				if(!isPLvlPropertiesAccepted){
					if(msg.isFullySuitable()){
						isPLvlPropertiesAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlPropertiesAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
			
			if(!result && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
				boolean isPLvlStyleAccepted = false; 
				msg = new UserMessage();
				msg = styleIdentifier.isSmallCapsAccepted(p, stylePart, predefinedStyle, msg);
				if(!isPLvlStyleAccepted){
					if(msg.isFullySuitable()){
						isPLvlStyleAccepted = true;
						result = true;
					} else {
						if(isDeepChk && msg.isChkresultSuitable()){
							isPLvlStyleAccepted = true;
							result = true;
						} else if(msg.isChkresultSuitable() && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
							isRLvlPropertiesAccepted = true;
							result = true;
						}
					}
				}
			}
		}
		
		return result;
	}
//	public boolean isFontSmallCapsSuitable(P p, Style predefinedStyle, StyleDefinitionsPart stylePart, boolean isDeepChk){
//		boolean result = false;
//		if(isDeepChk){
//			Style style = new Style();
//			style.setRPr(new RPr());
//			if(predefinedStyle.getRPr().getSmallCaps()!=null && predefinedStyle.getRPr().getSmallCaps().isVal()){
//				style.getRPr().setSmallCaps(new BooleanDefaultTrue());
//			} else {
//				BooleanDefaultTrue val = new BooleanDefaultTrue();
//				val.setVal(false);
//				style.getRPr().setSmallCaps(val);
//			}
//			result = chkStyleInR(p, style, stylePart);
//			if(result) return result;
//			else return false;
//		}		
//		if(predefinedStyle!=null){
//			Style chkingStyle = OOXMLConvertTool.create4ChkingStyle(new Style());
//			chkingStyle.getRPr().getSmallCaps().setVal(true);
//			ArrayList<String> appearedStyleList = new ArrayList<String>();
//			appearedStyleList.add(predefinedStyle.getStyleId());
//			boolean isShouldSmallCaps = rekursiveChkStyle(stylePart, chkingStyle, predefinedStyle, appearedStyleList, false);
//			if(p!=null && p.getPPr()!=null){
//				Style style = p.getPPr().getPStyle()!=null?stylePart.getStyleById(p.getPPr().getPStyle().getVal()):null;
//				Style rStyle = OOXMLConvertTool.getFirstRStyle(p);
//				if(rStyle != null){
//					style = rStyle;
//				}
//				if(p.getPPr().getRPr()!=null && p.getPPr().getRPr().getSmallCaps()!=null) {
//					if(p.getPPr().getRPr().getSmallCaps().isVal() && isShouldSmallCaps){
//						result = true;
//					}
//				} else if(style!=null){
//					appearedStyleList = new ArrayList<String>();
//					appearedStyleList.add(style.getStyleId());
//					result = rekursiveChkStyle(stylePart, chkingStyle, style, appearedStyleList, true) == isShouldSmallCaps;
//				}
//			} else {
//				result = true;
//			}
//		} else {
//			result = true;
//		}
//		
//		return result;
//	}
	
	public boolean isParagraphHasSectionNumber(P p, Style predefinedStyle, StyleDefinitionsPart stylePart){
		boolean result = false;
		
		if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
			if(predefinedStyle.getPPr()!=null && predefinedStyle.getPPr().getNumPr()!=null){
				result = true;
			} else {
				result = false;
			}
		} else if(p.getPPr()!=null && p.getPPr().getPStyle()!=null) {
			Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
			if(style!=null){
				if(style.getPPr()!=null && style.getPPr().getNumPr()!=null){
					if(predefinedStyle.getPPr()!=null && predefinedStyle.getPPr().getNumPr()!=null){
						result = true;
					} else {
						result = false;
					}
				} else {
					if(predefinedStyle.getPPr()!=null && predefinedStyle.getPPr().getNumPr()!=null){
						result = false;
					} else {
						result = true;
					}

				}
			}
		} else {
			if(predefinedStyle.getPPr()!=null && predefinedStyle.getPPr().getNumPr()!=null){
				result = false;
			} else {
				result = true;
			}
		}
		String pText = OOXMLConvertTool.getPText(p);
		
		return result;
	}
	
	public boolean isTitle(P p, StyleDefinitionsPart stylePart) {
		boolean isTitle = false;
		if(!isTitleIdentified){
			if(sourceFileFormat==null || sourceFileFormat!=null && sourceFileFormat.length()==0){
				for (Map.Entry<String, AcademicStyleMap> entry : AppEnvironment.getInstance().getStylePool().getStyles().entrySet()){
					AcademicStyleMap map = entry.getValue();
					if(map.getStyleMap().get(AcademicFormatStructureDefinition.TITLE)!=null){
						Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.TITLE);
						if(isFontSzSuitable(p, style, stylePart, true)
								&& isFontBoldSuitable(p, style, stylePart, true)){
							sourceFileFormat = entry.getKey();
							isTitle = true;
							isTitleIdentified = true;
						}
					}
				}
			} else {
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				if(map.getStyleMap().get(AcademicFormatStructureDefinition.TITLE)!=null){
					Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.TITLE);
					if(isFontSzSuitable(p, style, stylePart, false)
							&& isFontBoldSuitable(p, style, stylePart, true)){
						isTitle = true;
						isTitleIdentified = true;
					}
				}
			}
		} else {
			
		}
		return isTitle;
	}
	
	public boolean isSubtitle(P p, StyleDefinitionsPart stylePart){
		boolean isSubtitle = false;
		
		if(!isSubTitleIdentified){
			if(sourceFileFormat==null || sourceFileFormat!=null && sourceFileFormat.length()==0){
				for (Map.Entry<String, AcademicStyleMap> entry : AppEnvironment.getInstance().getStylePool().getStyles().entrySet()){
					AcademicStyleMap map = entry.getValue();
					if(map.getStyleMap().get(AcademicFormatStructureDefinition.SUBTITLE)!=null){
						Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.SUBTITLE);
						if(isFontSzSuitable(p, style, stylePart, false)){
							sourceFileFormat = entry.getKey();
							isSubtitle = true;
							isSubTitleIdentified = true;
						}
					}
				}
			} else {
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				if(map.getStyleMap().get(AcademicFormatStructureDefinition.SUBTITLE)!=null){
					Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.SUBTITLE);
					if(isFontSzSuitable(p, style, stylePart, false)){
						isSubtitle = true;
						isSubTitleIdentified = true;
					}
				}
			}
		} else {
			
		}
		
		return isSubtitle;
	}
	
	public boolean isAbstractHeader(P p, StyleDefinitionsPart stylePart){
		boolean isAbstract = false;
		
		if(!isAbstractIdentified){
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				if(map.getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACTHEADER)!=null){
					Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACTHEADER);
					String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
					String regex = "^abstract|^abstract\\..*|^abstract"+((char)8212)+".*|^abstract"+((char)8211)+".*";
					if(isFontSzSuitable(p, style, stylePart, true)
							&& isFontBoldSuitable(p, style, stylePart, true)
							&& isFontItalicSuitable(p, style, stylePart, true)
							&& pText.matches(regex)
							){
//					if(isFontSzSuitable(p, style, stylePart, false)
//							&& (OOXMLConvertTool.isPTextStartWithString(p, "Abstract", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Abstract")
//									|| OOXMLConvertTool.isPTextStartWithString(p, "Abstract.", true)
//									|| OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)8212), true)	// IEEE
//									|| OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)45), true))){
						isAbstract = true;
						isAbstractIdentified = true;
						isSubTitleIdentified = true;
					}
				}
			} else {
				for (Map.Entry<String, AcademicStyleMap> entry : AppEnvironment.getInstance().getStylePool().getStyles().entrySet()){
					AcademicStyleMap map = entry.getValue();
					if(map.getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACTHEADER)!=null){
						Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACTHEADER);
						String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
						String regex = "^abstract|^abstract\\..*|^abstract"+((char)8212)+".*|^abstract"+((char)8211)+".*";
						if(isFontSzSuitable(p, style, stylePart, true)
								&& isFontBoldSuitable(p, style, stylePart, true)
								&& isFontItalicSuitable(p, style, stylePart, true)
								&& pText.matches(regex)
								){
//						if(isFontSzSuitable(p, style, stylePart, false)
//								&& (OOXMLConvertTool.isPTextStartWithString(p, "Abstract", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Abstract")
//										|| OOXMLConvertTool.isPTextStartWithString(p, "Abstract.", true)
//										|| OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)8212), true)	// IEEE
//										|| OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)45), true))){
							sourceFileFormat = entry.getKey();
							isAbstract = true;
							isAbstractIdentified = true;
							isSubTitleIdentified = true;
						}
					}
				}
			}
		} 
		
		return isAbstract;
	}
	
	public boolean isKeywordHeader(P p, StyleDefinitionsPart stylePart){
		boolean isKeyword = false;
		if(!isKeywordIdentified){
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
				
//				pText.replaceAll(" ", "");
//				if(pText.toLowerCase().indexOf("keywords")>=0){
//					StringBuilder sbuilder = new StringBuilder(pText.substring(pText.indexOf("keywords")+"keywords".length()+1, pText.indexOf("keywords")+"keywords".length()+3));
//					char c = sbuilder.charAt(1);
//					System.out.println((char)8212);
//				}
				
				if(map.getStyleMap().get(AcademicFormatStructureDefinition.KEYWORDHEADER)!=null && pText!=null && pText.replaceAll(" ", "").length()>0){
					Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.KEYWORDHEADER);
					String regex = "^keywords|^keywords:.*|^keywords"+((char)8212)+".*|^keywords"+((char)8211)+".*|^keywords-.*";
					if(isFontSzSuitable(p, style, stylePart, true)
							&& isFontBoldSuitable(p, style, stylePart, true)
							&& isFontItalicSuitable(p, style, stylePart, true)
							&& pText.matches(regex)
							){
//					if(isFontSzSuitable(p, style, stylePart, false)
//							&& (OOXMLConvertTool.isPTextStartWithString(p, "Keywords", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Abstract")
//									|| OOXMLConvertTool.isPTextStartWithString(p, "Keywords:", true)
//									|| OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)8212), true)	// IEEE
//									|| OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)45), true))){
						isKeyword = true;
						isKeywordIdentified = true;
					}
				}
			} else {
				for (Map.Entry<String, AcademicStyleMap> entry : AppEnvironment.getInstance().getStylePool().getStyles().entrySet()){
					AcademicStyleMap map = entry.getValue();
					if(map.getStyleMap().get(AcademicFormatStructureDefinition.KEYWORDHEADER)!=null){
						Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.KEYWORDHEADER);
						String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
						String regex = "^keywords|^keywords:.*|^keywords"+((char)8212)+".*|^keywords"+((char)8211)+".*";
						if(isFontSzSuitable(p, style, stylePart, false)
								&& isFontBoldSuitable(p, style, stylePart, true)
								&& isFontItalicSuitable(p, style, stylePart, true)
								&& pText.matches(regex)
								){
//						if(isFontSzSuitable(p, style, stylePart, false)
//								&& (OOXMLConvertTool.isPTextStartWithString(p, "Keywords", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Abstract")
//										|| OOXMLConvertTool.isPTextStartWithString(p, "Keywords:", true)
//										|| OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)8212), true)	// IEEE
//										|| OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)45), true))){
							sourceFileFormat = entry.getKey();
							isKeyword = true;
							isKeywordIdentified = true;
						}
					}
				}
			}
		} 
		
		return isKeyword;
	}
	
	public boolean isAcknowledgmentHeader(P p, NumberingDefinitionsPart numberingPart, StyleDefinitionsPart stylePart){
		boolean isAcknowledgment = false;
		if(!isAcknowledgeIdentified){
			P newP = OOXMLConvertTool.cloneP(p);
			OOXMLConvertTool.removeSectionNumStringOfHeading(newP);
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				if(map.getStyleMap().get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER)!=null){
					Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER);
					String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
					String regex = "(^acknowledgment|^acknowledgments|^acknowledgments\\.).*";
					if(isFontSzSuitable(newP, style, stylePart, false)
							&& isFontBoldSuitable(newP, style, stylePart, true)
							&& isFontItalicSuitable(newP, style, stylePart, true)
							&& isFontCapsSuitable(newP, style, stylePart, true)
							&& isFontSmallCapsSuitable(newP, style, stylePart, true)
							&& pText.matches(regex)){
//							&& (OOXMLConvertTool.isPTextStartWithString(newP, "Acknowledgments", true) && OOXMLConvertTool.getPText(newP).replaceAll(" ", "").equals("Acknowledgments")	// ACM and Elsevier
//									|| OOXMLConvertTool.isPTextStartWithString(newP, "Acknowledgment", true) && OOXMLConvertTool.getPText(newP).replaceAll(" ", "").equals("Acknowledgment")	// IEEE
//									|| OOXMLConvertTool.isPTextStartWithString(newP, "Acknowledgments.", true))){	// Springer
						isAcknowledgment = true;
						isAcknowledgeIdentified = true;
					}
				}
			} else {
				for (Map.Entry<String, AcademicStyleMap> entry : AppEnvironment.getInstance().getStylePool().getStyles().entrySet()){
					AcademicStyleMap map = entry.getValue();
					if(map.getStyleMap().get(AcademicFormatStructureDefinition.KEYWORDHEADER)!=null){
						Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.KEYWORDHEADER);
						String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
						String regex = "(^acknowledgment|^acknowledgments|^acknowledgments\\.).*";
						if(isFontSzSuitable(newP, style, stylePart, false)
								&& isFontBoldSuitable(newP, style, stylePart, true)
								&& isFontItalicSuitable(newP, style, stylePart, true)
								&& isFontCapsSuitable(newP, style, stylePart, true)
								&& isFontSmallCapsSuitable(newP, style, stylePart, true)
								&& pText.matches(regex)){
//						if(isFontSzSuitable(newP, style, stylePart, false)
//								&& (OOXMLConvertTool.isPTextStartWithString(newP, "Acknowledgments", true) && OOXMLConvertTool.getPText(newP).replaceAll(" ", "").equals("Acknowledgments")	// ACM and Elsevier
//										|| OOXMLConvertTool.isPTextStartWithString(newP, "Acknowledgment", true) && OOXMLConvertTool.getPText(newP).replaceAll(" ", "").equals("Acknowledgment")	// IEEE
//										|| OOXMLConvertTool.isPTextStartWithString(newP, "Acknowledgments.", true))){	// Springer
							sourceFileFormat = entry.getKey();
							isAcknowledgment = true;
							isAcknowledgeIdentified = true;
						}
					}
				}
			}
		}
		
		return isAcknowledgment;
	}
	
	public boolean isReferenceHeader(P p, NumberingDefinitionsPart numberingPart, StyleDefinitionsPart stylePart){
		boolean isReferenceHeader = false;
		P newP = OOXMLConvertTool.cloneP(p);
		OOXMLConvertTool.removeSectionNumStringOfHeading(newP);
		if(!isReferencesHeaderIdentified){
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				if(map.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER)!=null){
					Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER);
					String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
					String regex = "^references";
					if(isFontSzSuitable(newP, style, stylePart, false)
							&& isFontBoldSuitable(newP, style, stylePart, true)
							&& isFontItalicSuitable(newP, style, stylePart, true)
							&& isFontCapsSuitable(newP, style, stylePart, false)
							&& isFontSmallCapsSuitable(newP, style, stylePart, false)
							&& pText.matches(regex)){
//					if(isFontSzSuitable(newP, style, stylePart, false) 
//							&& isFontCapsSuitable(newP, style, stylePart, false)
//							&& isFontBoldSuitable(newP, style, stylePart, false)
//							&& isFontItalicSuitable(newP, style, stylePart, false)
//							&& isFontSmallCapsSuitable(newP, style, stylePart, false)
//							&& (OOXMLConvertTool.isPTextStartWithString(newP, "References", false) && OOXMLConvertTool.getPText(newP).replaceAll(" ", "").toLowerCase().equals("References".toLowerCase()))) {
						isReferenceHeader = true;
						isReferencesHeaderIdentified = true;
					}
				}
			} else {
				for (Map.Entry<String, AcademicStyleMap> entry : AppEnvironment.getInstance().getStylePool().getStyles().entrySet()){
					AcademicStyleMap map = entry.getValue();
					if(map.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER)!=null){
						Style style = map.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER);
						String pText = OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "");
						String regex = "^references";
						if(isFontSzSuitable(newP, style, stylePart, false)
								&& isFontBoldSuitable(newP, style, stylePart, true)
								&& isFontItalicSuitable(newP, style, stylePart, true)
								&& isFontCapsSuitable(newP, style, stylePart, false)
								&& isFontSmallCapsSuitable(newP, style, stylePart, false)
								&& pText.matches(regex)){
//						if(isFontSzSuitable(newP, style, stylePart, false)
//								&& (OOXMLConvertTool.isPTextStartWithString(newP, "References", false) && OOXMLConvertTool.getPText(newP).replaceAll(" ", "").toLowerCase().equals("References".toLowerCase()))) {
							isReferenceHeader = true;
							isReferencesHeaderIdentified = true;
						}
					}
				}
			}
		}
		
		return isReferenceHeader;
	}
	
	private boolean isB(Style chkingStyle, boolean isB){
		boolean result = false;
		if(chkingStyle!=null){
			if(chkingStyle.getRPr()!=null){
				if(chkingStyle.getRPr().getB()!=null){
					if(chkingStyle.getRPr().getB().isVal()&&isB){
						result = true;
					} else {
						result = false;
					}
				} else if(!isB) {
					result = true;
				} else {
					result = false;
				}
			}
		}
		return result;
	}
	public boolean rekursiveChkBInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg){
		boolean result = false;
		if(style!=null){
			if(style.getRPr()!=null && style.getRPr().getB()!=null){
				if(style.getRPr().getB().isVal()){
					result = true;
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
				} else {
					result = false;
					msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
				}
			} else {
				Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
				result = rekursiveChkBInStyle(basedOnStyle, stylePart, msg);
				if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
					result = false;
				} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
					Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
					result = rekursiveChkBInStyle(linkStyle, stylePart, msg);
					if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						result = false;
					} else {
						result = true;
					}
				} else {
					result = true;
				}
			}
		}
		return result;
	}
	public boolean rekursiveChkIInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg){
		boolean result = false;
		if(style!=null){
			if(style.getRPr()!=null && style.getRPr().getI()!=null){
				if(style.getRPr().getI().isVal()){
					result = true;
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
				} else {
					result = false;
					msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
				}
			} else {
				Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
				result = rekursiveChkBInStyle(basedOnStyle, stylePart, msg);
				if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
					result = false;
				} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
					Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
					result = rekursiveChkBInStyle(linkStyle, stylePart, msg);
					if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						result = false;
					} else {
						result = true;
					}
				} else {
					result = true;
				}
			}
		}
		return result;
	}
	public boolean isPStyleSuitable(Style chkingStyle, StyleDefinitionsPart stylePart, Style predefinedStyle){
		boolean result = false;
		UserMessage msg = new UserMessage();
		if(chkingStyle!=null){
			if(chkingStyle.getRPr()!=null){
				boolean isShouldB = chkingStyle.getRPr().getB()!=null?chkingStyle.getRPr().getB().isVal():false;
				
			} else {
				
			}
		} else {
			
			return false;
		}
		return result;
	}
	
	public boolean isHeadingProperites(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, boolean isDeep){
		boolean isHeading = false;
		
		if(p!=null){
			/*
			 * There is numPr, -> chk properties
			 */
			if(isFontBoldSuitable(p, predefinedStyle, stylePart, isDeep)
					&& isFontCapsSuitable(p, predefinedStyle, stylePart, isDeep)
					&& isFontItalicSuitable(p, predefinedStyle, stylePart, isDeep)
					&& isFontSmallCapsSuitable(p, predefinedStyle, stylePart, isDeep)
					&& isFontSzSuitable(p, predefinedStyle, stylePart, isDeep)
					&& isParagraphHasSectionNumber(p, predefinedStyle, stylePart)
					){
				isHeading = true;
			}
		}
		
		return isHeading;
	}
	
	public int getHeadingLvl(P p, NumberingDefinitionsPart numberingPart, StyleDefinitionsPart stylePart){
		int headingLvl = 0;
		
		if(isReferencesHeaderIdentified){
			return headingLvl;
		}
		
		String[] regex = {"(^(\\d+)(\\.){0,1}(["+((char)32)+"|"+((char)8195)+"]{0,1}).*$)|(^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})(\\.["+((char)32)+"|"+((char)8195)+"]{1,1}).*$)",
				"(^(\\d+)(\\.\\d+)(\\.){0,1}(["+((char)9)+"|"+((char)32)+"|"+((char)8195)+"]{0,1}).*$)|(^([A-Z]+\\.)(["+((char)32)+"|"+((char)8195)+"]{1,1}).*)",
				"(^(\\d+\\.\\d+\\.\\d+)(\\.){0,1}(["+((char)32)+"|"+((char)8195)+"]{0,1}).*$)|(^(\\d+\\))(["+((char)32)+"|"+((char)8195)+"]{1,1}).*)",
				"(^(\\d+)(\\.\\d+){3,3}(\\.){0,1}(["+((char)32)+"|"+((char)8195)+"]{0,1}).*$)|(^([a-z]+\\))(["+((char)32)+"|"+((char)8195)+"]{1,1}).*)"
		};
		
		
		try {
			if(p!=null){
				/*
				 * chk if the section number would be generated manually, and try to get the section level.
				 */
				String pText = OOXMLConvertTool.getPText(p);
				for (int i = 0; i < regex.length; i++) {
					if(pText.matches(regex[i])) {
						headingLvl = i+1;
						break;
					} else if(p.toString().matches(regex[i])){
						headingLvl = i+1;
						break;
					}
				}
				/*
				 * chk if the section number would be generated automatically
				 */
				if(headingLvl == 0){
					NumPr numPr = null;
					if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
						/*
						 * chk in document.xml p-tag property
						 */
						numPr = p.getPPr().getNumPr();
					} else if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
						/*
						 * chk style definition
						 */
						Style pStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
						if(pStyle!=null && pStyle.getPPr()!=null){
							numPr = pStyle.getPPr().getNumPr();
						}
					}
					
					if(numPr!=null){
						AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
						String styleId = AcademicFormatStructureDefinition.HEADING1;
						styleId = styleId.substring(0, styleId.length()-1);
						for(int i=1; i<=4; i++){
							Style predefinedStyle = map.getStyleMap().get(styleId+i);
							OOXMLIdentifierProvider idProvider = new OOXMLIdentifierProvider(sourceFileFormat);
							UserMessage msg = new UserMessage();
							if(idProvider.isNumPrAcceptable(p, i-1, stylePart, numberingPart, msg)
									&& isHeadingProperites(p, stylePart, predefinedStyle, true)){
								headingLvl = i;
								break;
							}
						}
					} else {
						if(!OOXMLConvertTool.isAllRInSameStyle(p, stylePart)){
							AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
							String styleId = AcademicFormatStructureDefinition.HEADING1;
							styleId = styleId.substring(0, styleId.length()-1);
							for(int i=1; i<=4; i++){
								Style predefinedStyle = map.getStyleMap().get(styleId+i);
								if(isHeadingProperites(p, stylePart, predefinedStyle, true)){
									headingLvl = i;
									break;
								}
							}
						} else {
							
						}
					}
				}
			}	
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return headingLvl;
	}
	private AbstractNum getAbstractNum(List<AbstractNum> abstractNumList, List<Num> numList, String numId) {
		AbstractNum abstractNum = null;
		BigInteger abNumId = new BigInteger(numId);
		
		try {
			for (int i = 0; i < numList.size(); i++) {
				Num num = numList.get(i);
				if(num.getNumId()!=null && num.getNumId().toString().equals(numId)){
					abNumId = num.getAbstractNumId().getVal();
					break;
				}
			}
			
			for(int index=0; index<abstractNumList.size(); index++) {
				AbstractNum temp = abstractNumList.get(index);
				if (temp.getAbstractNumId().intValue() == abNumId.intValue()) {
					abstractNum = temp;
					break;
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return abstractNum;
	}
	private boolean isCaptionNumbered(P p, StyleDefinitionsPart stylePart, NumberingDefinitionsPart numberingPart, boolean isTable) {
		boolean isNumbered = false;
		
		try {
			Numbering numbering = numberingPart.getContents();
			List<AbstractNum> abNumList = numbering.getAbstractNum();
			List<Num> numList = numbering.getNum();
			
			String numId = "";
			if(p.getPPr()!=null){
				if(p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal().toString();
					}
				} else {
					if(p.getPPr().getPStyle()!=null){
						Style style = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
						if(style!=null && style.getPPr()!=null && style.getPPr().getNumPr()!=null){
							numId = style.getPPr().getNumPr().getNumId().getVal().toString();
						}
					}
				}
			}
			if(numId!=null && numId.length()>0){
				AbstractNum abstNum = getAbstractNum(abNumList, numList, numId);
				if(abstNum!=null && abstNum.getLvl()!=null && abstNum.getLvl().get(0)!=null){
					if(isTable){
						if(abstNum.getLvl().get(0).getLvlText()!=null && "table%1.".equals(abstNum.getLvl().get(0).getLvlText().getVal().replaceAll(" ", "").toLowerCase())){
							isNumbered = true;
						}
					} else {
						if(abstNum.getLvl().get(0).getLvlText()!=null && "figure%1.".equals(abstNum.getLvl().get(0).getLvlText().getVal().replaceAll(" ", "").toLowerCase())){
							isNumbered = true;
						} else if(abstNum.getLvl().get(0).getLvlText()!=null && "fig.%1.".equals(abstNum.getLvl().get(0).getLvlText().getVal().replaceAll(" ", "").toLowerCase())){
							isNumbered = true;
						}
					}
					
				}
			}
			
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			LogUtils.log(e.getMessage());
		}
		
		
		return isNumbered;
	}
	
	public boolean isFigureCaption(P p, StyleDefinitionsPart stylePart, NumberingDefinitionsPart numberingPart){
		boolean isCaption = false;
		
		try {
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				String pattern = "^Figure \\d+.*|^Fig. \\d+.*";
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				Style predefinedStyle = map.getStyleMap().get(AcademicFormatStructureDefinition.CAPTIONHEADER);
				String pText = getPText(p);
				if(pText==null || pText!=null && pText.length()<=0){
					return isCaption;
				}
				if(isFontBoldSuitable(p, predefinedStyle, stylePart, true)
						&& isFontSzSuitable(p, predefinedStyle, stylePart, false)
						&& isFontItalicSuitable(p, predefinedStyle, stylePart, false)
						&& isFontCapsSuitable(p, predefinedStyle, stylePart, true)
						&& isFontSmallCapsSuitable(p, predefinedStyle, stylePart, false)
						&& pText!=null && (pText.matches(pattern)|| isCaptionNumbered(p, stylePart, numberingPart, false))){
					isCaption = true;
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return isCaption;
	}
	
	public boolean isTableCaption(P p, StyleDefinitionsPart stylePart, NumberingDefinitionsPart numberingPart){
		boolean isCaption = false;
		try {
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				String pattern = "^Table \\d+.*";
				AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
				Style predefinedStyle = map.getStyleMap().get(AcademicFormatStructureDefinition.CAPTIONHEADER);
				String pText = getPText(p);
				if(pText==null || pText!=null && pText.length()<=0){
					return isCaption;
				}
				if(isFontBoldSuitable(p, predefinedStyle, stylePart, true)
						&& isFontSzSuitable(p, predefinedStyle, stylePart, false)
						&& isFontItalicSuitable(p, predefinedStyle, stylePart, false)
						&& isFontCapsSuitable(p, predefinedStyle, stylePart, true)
						&& isFontSmallCapsSuitable(p, predefinedStyle, stylePart, false)
						&& pText!=null && (pText.matches(pattern)|| isCaptionNumbered(p, stylePart, numberingPart, true))){
					isCaption = true;
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return isCaption;
	}
	
	public boolean isAppendix(P p, StyleDefinitionsPart stylePart) {
		boolean isAppendix = false;
		try {
			if(sourceFileFormat!=null && sourceFileFormat.length()>0){
				String pText = getPText(p);
				if(pText!=null && pText.replaceAll(" ", "").length()>0){
					AcademicStyleMap map = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFileFormat);
					Style predefinedStyle = map.getStyleMap().get(AcademicFormatStructureDefinition.APPENDIXHEADER);
					if(isFontBoldSuitable(p, predefinedStyle, stylePart, false)
							&& isFontSzSuitable(p, predefinedStyle, stylePart, false)
							&& isFontItalicSuitable(p, predefinedStyle, stylePart, false)
							&& isFontCapsSuitable(p, predefinedStyle, stylePart, true)
							&& isFontSmallCapsSuitable(p, predefinedStyle, stylePart, false)
							&& pText.toLowerCase().startsWith("appendix")){
						isAppendix = true;
						isAppendixHeaderIdentified = true;
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return isAppendix;
	}
}
