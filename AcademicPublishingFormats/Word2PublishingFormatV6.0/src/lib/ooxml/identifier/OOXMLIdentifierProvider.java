package lib.ooxml.identifier;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.P;
import org.docx4j.wml.Style;
import org.docx4j.wml.DocDefaults.PPrDefault;
import org.docx4j.wml.DocDefaults.RPrDefault;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.Num;
import org.docx4j.wml.Numbering.Num.AbstractNumId;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import tools.LogUtils;

public class OOXMLIdentifierProvider {
	private String sourceFormat;
	public OOXMLIdentifierProvider(String sourceFormat) {
		// TODO Auto-generated constructor stub
		this.sourceFormat = sourceFormat;
	}
	public String getSourceFormat() {
		return sourceFormat;
	}
	public void setSourceFormat(String sourceFormat) {
		this.sourceFormat = sourceFormat;
	}
	public boolean rekursiveChkBInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getB()!=null){
					if(style.getRPr().getB().isVal()){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
					} else {
						result = false;
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkBInStyle(basedOnStyle, stylePart, msg, appearedStyleList, isFromDoc);
					} else {
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
					if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						result = false;
					} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
						Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
						if(linkStyle!=null && linkStyle.getType().equals("paragraph")){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkBInStyle(linkStyle, stylePart, msg, appearedStyleList, isFromDoc);
						}
						if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
							result = false;
						} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
							result = false;
						} else {
							result = true;
						}
					} else {
						result = true;
					}
				}
			} else {
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	public boolean rekursiveChkIInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getI()!=null){
					if(style.getRPr().getI().isVal()){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkIInStyle(basedOnStyle, stylePart, msg, appearedStyleList, isFromDoc);
					}
					if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
						Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkIInStyle(linkStyle, stylePart, msg, appearedStyleList, isFromDoc);
						}
						if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
							result = false;
						} else {
							result = false;
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						}
					} else {
						result = true;
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
					}
				}
			} else {
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	public boolean rekursiveChkCapsInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getCaps()!=null){
					if(style.getRPr().getCaps().isVal()){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkCapsInStyle(basedOnStyle, stylePart, msg, appearedStyleList, isFromDoc);
					}
					if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						result = false;
					} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
						Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkCapsInStyle(linkStyle, stylePart, msg, appearedStyleList, isFromDoc);
						}
						if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
							result = false;
						} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
							result = false;
						} else {
							result = true;
						}
					} else {
						result = true;
					}
				}
			} else {
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
				//LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	public boolean rekursiveChkSmallCapsInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getSmallCaps()!=null){
					if(style.getRPr().getSmallCaps().isVal()){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkSmallCapsInStyle(basedOnStyle, stylePart, msg, appearedStyleList, isFromDoc);
					}
					
					if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						result = false;
					} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())) {
						Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkSmallCapsInStyle(basedOnStyle, stylePart, msg, appearedStyleList, isFromDoc);
						}
						if(msg!=null && UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
							result = false;
						} else if(msg!=null && UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
							result = false;
						} else {
							result = true;
						}
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
				}
			} else {
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	public BigDecimal rekursiveGetSzInStyle(Style style, StyleDefinitionsPart stylePart, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		BigDecimal result = null;
		try {
			if(style!=null){
				if(style.getRPr()!=null && style.getRPr().getSz()!=null && style.getRPr().getSz().getVal()!=null){
					result = new BigDecimal(style.getRPr().getSz().getVal().intValue());
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? stylePart.getStyleById(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = null;
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveGetSzInStyle(basedOnStyle, stylePart, msg, appearedStyleList, isFromDoc);
					} else {
						result = null;
					}
					if(result==null){
						Style linkStyle = style.getLink()!=null ? stylePart.getStyleById(style.getLink().getVal()) : null;
						if(linkStyle!=null && linkStyle.getType().equals("paragraph")){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = null;
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveGetSzInStyle(linkStyle, stylePart, msg, appearedStyleList, isFromDoc);
						} else {
							result = null;
						}
					} else {
						
					}
				}
			} else {
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
			}
			if(result==null){
				result = new BigDecimal("20");
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	private BigInteger getBasedOnStyleNumId(Style style, StyleDefinitionsPart stylePart){
		BigInteger result = null;
		UserMessage msg = new UserMessage();
		try {
			Style basedOnStyle = null;
			if(style.getBasedOn()!=null){
				basedOnStyle = stylePart.getStyleById(style.getBasedOn().getVal());
				if(basedOnStyle.getPPr()!=null && basedOnStyle.getPPr().getNumPr()!=null && basedOnStyle.getPPr().getNumPr().getNumId()!=null){
					result = basedOnStyle.getPPr().getNumPr().getNumId().getVal();
				} else {
					result = getBasedOnStyleNumId(basedOnStyle, stylePart);
				}
			} else {
				// return null
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	public boolean isNumPrAcceptable(P p, int lvlIndex, StyleDefinitionsPart stylePart, NumberingDefinitionsPart numberingPart, UserMessage msg){
		boolean result = false;
		try {
			if(p!=null){
				BigInteger numId = null;
				BigInteger lvl = null;
				if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal();
					}
					if(p.getPPr().getNumPr().getIlvl()!=null){
						lvl = p.getPPr().getNumPr().getIlvl().getVal();
					}
				} else if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
					Style pStyle = stylePart.getStyleById(p.getPPr().getPStyle().getVal());
					if(pStyle.getPPr()!=null && pStyle.getPPr().getNumPr()!=null){
						if(pStyle.getPPr().getNumPr().getNumId()!=null){
							numId = pStyle.getPPr().getNumPr().getNumId().getVal();
						} else {
							numId = getBasedOnStyleNumId(pStyle, stylePart);
						}
						if(pStyle.getPPr().getNumPr().getIlvl()!=null){
							lvl = pStyle.getPPr().getNumPr().getIlvl().getVal();
						}
					}
				}
				if(numId!=null){
					try {
						List<Num> numList = numberingPart.getContents().getNum();
						AbstractNumId absNumId = null;
						for(int x=0; x<numList.size(); x++){
							Num n = numList.get(x);
							if(n.getNumId().intValue()==numId.intValue()){
								absNumId = n.getAbstractNumId();
								break;
							}
						}
						if(absNumId==null){
							if(lvl.intValue()==0){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								return result;
							}
							absNumId = new AbstractNumId();
							absNumId.setVal(numId);
						}
						List<AbstractNum> abstNumList = numberingPart.getContents().getAbstractNum();
						AbstractNum abstNumDef = AppEnvironment.getInstance().getStylePool().getStyles().get(this.sourceFormat).getNumberingMap().get(AcademicFormatStructureDefinition.NUMBERINGSECTIONSTRING);
						for(int i =0; i<abstNumDef.getLvl().size(); i++){
							if(abstNumDef.getLvl().get(i).getIlvl().intValue()==lvlIndex){
								lvlIndex = i;
								break;
							}
						}
						
						ROOT: for (int i = 0; i < abstNumList.size(); i++) {
							AbstractNum abstNum = abstNumList.get(i);
							if(abstNum.getAbstractNumId()!=null && absNumId != null && absNumId.getVal()!=null && abstNum.getAbstractNumId().intValue() == absNumId.getVal().intValue()){
								List<Lvl> lvlList = abstNum.getLvl();
								if(lvl==null){
									lvl = BigInteger.valueOf(0);
								}
								
								for (int j = 0; j < lvlList.size(); j++) {
									boolean chk1= false, chk2= false;
									if(lvlList.get(j).getIlvl().intValue()==lvl.intValue()){
										if(abstNumDef!=null && abstNumDef.getLvl()!=null && abstNumDef.getLvl().get(lvlIndex)!=null){
											if(lvlList.get(j).getLvlText()!=null
													&& abstNumDef.getLvl().get(lvlIndex).getLvlText()!=null){
												if(lvlList.get(j).getLvlText().getVal().equals(abstNumDef.getLvl().get(lvlIndex).getLvlText().getVal())){
													chk1 = true;
												}
											}
											if(lvlList.get(j).getNumFmt()!=null && abstNumDef.getLvl().get(lvlIndex).getNumFmt()!=null){
												if(abstNumDef.getLvl().get(lvlIndex).getNumFmt().getVal().equals(lvlList.get(j).getNumFmt().getVal())){
													chk2 = true;
												}
											}
											if(chk1&&chk2){
												result = true;
												msg.setFullySuitable(true);
												msg.setChkresultSuitable(true);
												msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
												break ROOT;
											} else {
												msg.setFullySuitable(false);
												msg.setChkresultSuitable(false);
												msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
												continue;
											}
										}
									}
								}
							}
						}
					} catch (Docx4JException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return result;
	}
	public boolean rekursiveChkStyle(Style style, StyleDefinitionsPart stylePart, boolean isB, boolean isI, boolean isCaps, boolean isSmallCaps){
		boolean result = false;
		boolean resB, resI, resCaps, resSmallCaps;
		UserMessage msg2 = new UserMessage();
		try {
			if(isB){
				ArrayList<String> appearedStyleList = new ArrayList<String>();
				appearedStyleList.add(style.getStyleId());
				UserMessage msg = new UserMessage();
				resB = rekursiveChkBInStyle(style, stylePart, msg, appearedStyleList, true);
				if(msg.isChkresultSuitable() && !UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
					resB = true;
				} else if(msg.isFullySuitable()){
					resB = true;
				} else {
					resB = false;
				}
			} else {
				resB = true;
			}
			
			if(isI){
				ArrayList<String> appearedStyleList = new ArrayList<String>();
				appearedStyleList.add(style.getStyleId());
				UserMessage msg = new UserMessage();
				resI = rekursiveChkIInStyle(style, stylePart, msg, appearedStyleList, true);
				if(msg.isChkresultSuitable() && !UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
					resI = true;
				} else if(msg.isFullySuitable()){
					resI = true;
				} else {
					resI = false;
				}
			} else {
				resI = true;
			}
			
			if(isCaps){
				ArrayList<String> appearedStyleList = new ArrayList<String>();
				appearedStyleList.add(style.getStyleId());
				UserMessage msg = new UserMessage();
				resCaps = rekursiveChkCapsInStyle(style, stylePart, msg, appearedStyleList, true);
				if(msg.isChkresultSuitable() && !UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
					resCaps = true;
				} else if(msg.isFullySuitable()){
					resCaps = true;
				} else {
					resCaps = false;
				}
			} else {
				resCaps = true;
			}

			if(isSmallCaps){
				ArrayList<String> appearedStyleList = new ArrayList<String>();
				appearedStyleList.add(style.getStyleId());
				UserMessage msg = new UserMessage();
				resSmallCaps = rekursiveChkSmallCapsInStyle(style, stylePart, msg, appearedStyleList, true);
				if(msg.isChkresultSuitable() && !UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
					resSmallCaps = true;
				} else if(msg.isFullySuitable()){
					resSmallCaps = true;
				} else {
					resSmallCaps = false;
				}
			} else {
				resSmallCaps = true;
			}
			
			if(resB && resI && resCaps && resSmallCaps) {
				result = true;
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg2.setMessageDetails(e.getMessage());
			LogUtils.log(msg2.getMessage());
		}
		
		return result;
	}
	
	public PPrDefault getDocDefaultPPr(StyleDefinitionsPart stylePart){
		PPrDefault pPr = null;
			try {
				List<Object> jaxbNodes = stylePart.getJAXBNodesViaXPath("w:docDefaults", true);
				if(jaxbNodes!=null){
					for (int i = 0; i < jaxbNodes.size(); i++) {
						if(jaxbNodes.get(i) instanceof DocDefaults){
							pPr = ((DocDefaults) jaxbNodes.get(i)).getPPrDefault();
						}
					}
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		
		return pPr;
	}
	
	public RPrDefault getDocDefaultRPr(StyleDefinitionsPart stylePart){
		RPrDefault rPr = null;
			try {
				List<Object> jaxbNodes = stylePart.getJAXBNodesViaXPath("w:docDefaults", false);
				if(jaxbNodes!=null){
					for (int i = 0; i < jaxbNodes.size(); i++) {
						if(jaxbNodes.get(i) instanceof DocDefaults){
							rPr = ((DocDefaults) jaxbNodes.get(i)).getRPrDefault();
						}
					}
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		
		return rPr;
	}
}
