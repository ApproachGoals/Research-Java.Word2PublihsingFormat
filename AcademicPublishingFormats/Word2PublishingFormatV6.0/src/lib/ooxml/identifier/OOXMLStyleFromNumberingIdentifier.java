package lib.ooxml.identifier;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.Num;
import org.docx4j.wml.P;
import org.docx4j.wml.Style;

import base.AppEnvironment;
import base.UserMessage;
import tools.LogUtils;

public class OOXMLStyleFromNumberingIdentifier extends OOXMLIdentifierProvider implements Identifier {

	public OOXMLStyleFromNumberingIdentifier(String sourceFormat) {
		super(sourceFormat);
		// TODO Auto-generated constructor stub
	}

	private String sourceFormat;
	
	public String getSourceFormat() {
		return sourceFormat;
	}
	public void setSourceFormat(String sourceFormat) {
		this.sourceFormat = sourceFormat;
	}

	
	private BigInteger recursiveAbsNumLvl(Style style, StyleDefinitionsPart stylePart, int mode, ArrayList<String> appearedStyleList){
		BigInteger numbId = null;
		BigInteger lvl = null;
		try {
			if(style!=null) {
				if(style.getPPr()!=null && style.getPPr().getNumPr()!=null){
					if(style.getPPr().getNumPr().getNumId()!=null){
						numbId = style.getPPr().getNumPr().getNumId().getVal();
					}
					if(style.getPPr().getNumPr().getIlvl()!=null){
						lvl = style.getPPr().getNumPr().getIlvl().getVal();
					}
				} else if(style.getBasedOn()!=null){
					Style basedonStyle = stylePart.getStyleById(style.getBasedOn().getVal());
					if(basedonStyle.getPPr()!=null && basedonStyle.getPPr().getNumPr()!=null){
						if(appearedStyleList.indexOf(style.getBasedOn().getVal())>=0){
							return null;
						}
						appearedStyleList.add(style.getBasedOn().getVal());
						if(mode==0){
							numbId = recursiveAbsNumLvl(basedonStyle, stylePart, mode, appearedStyleList);
						} else {
							lvl = recursiveAbsNumLvl(basedonStyle, stylePart, mode, appearedStyleList);
						}
					}
				} else if(style.getLink()!=null){
					Style linkStyle = stylePart.getStyleById(style.getLink().getVal());
					if(linkStyle.getPPr()!=null && linkStyle.getPPr().getNumPr()!=null){
						if(appearedStyleList.indexOf(style.getLink().getVal())>=0){
							return null;
						}
						appearedStyleList.add(style.getLink().getVal());
						if(mode==0){
							numbId = recursiveAbsNumLvl(linkStyle, stylePart, mode, appearedStyleList);
						} else {
							lvl = recursiveAbsNumLvl(linkStyle, stylePart, mode, appearedStyleList);
						}
					}
				}
			}
			
			
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		if(mode==0){
			return numbId;
		} else {
			return lvl;
		}
		
	}
	
	private AbstractNum getAbstNum(BigInteger numbId){
		AbstractNum absNum = null;
		
		NumberingDefinitionsPart numberingPart = AppEnvironment.getInstance().getNumberingPart();
		
		if(numbId!=null){
			try {
				ArrayList<Num> numList = (ArrayList<Num>) numberingPart.getContents().getNum();
				for(int i=0; i<numList.size(); i++){
					Num n = numList.get(i);
					if(n.getNumId()!=null && n.getNumId().intValue() == numbId.intValue()){
						numbId = new BigInteger("0");
						numbId = n.getAbstractNumId().getVal();
						break;
					}
				}
				
				ArrayList<AbstractNum> absNumList = (ArrayList<AbstractNum>) numberingPart.getContents().getAbstractNum();
				for(int i=0; i<absNumList.size(); i++){
					if(absNumList.get(i).getAbstractNumId().intValue()==numbId.intValue()){
						absNum = absNumList.get(i);
						break;
					}
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		}
		
		return absNum;
	}
	/*
	 * B
	 */
	public boolean rekursiveChkBInStyle(Style style, ArrayList<String> appearedStyleList){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getB()!=null){
					if(style.getRPr().getB().isVal()){
						result = true;
					} else {
						result = false;
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkBInStyle(basedOnStyle, appearedStyleList);
					}
					if(!result) {
						Style linkStyle = style.getLink()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkBInStyle(linkStyle, appearedStyleList);
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
	private boolean isStyleB(Style style){
		ArrayList<String> appearedStyleList = new ArrayList<String>();
		appearedStyleList.add(style.getStyleId());
		return rekursiveChkBInStyle(style, appearedStyleList);
	}
	public UserMessage isFontBAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		
		BigInteger numId = null;
		BigInteger lvl = null;
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal();
					}
					if(p.getPPr().getNumPr().getIlvl()!=null){
						lvl = p.getPPr().getNumPr().getIlvl().getVal();
					}
				} else {
					if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
						Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(p.getPPr().getPStyle().getVal());
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						numId = recursiveAbsNumLvl(style, stylePart, 0, appearedStyleList);
						appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						lvl = recursiveAbsNumLvl(style, stylePart, 1, appearedStyleList);
					}
				}
				if(numId!=null){
					AbstractNum absNum = getAbstNum(numId);
					if(absNum!=null){
						if(lvl==null){
							lvl = new BigInteger("1");
						}
						List<Lvl> lvlList = absNum.getLvl();
						for(int i=0; i<lvlList.size(); i++){
							if(lvlList.get(i)!=null && lvlList.get(i).getIlvl()!=null && lvlList.get(i).getIlvl().intValue() == lvl.intValue()){
								Lvl lvlData = lvlList.get(i);
								if(lvlData.getRPr()!=null){
									if(isStyleB(predefinedStyle)){
										if(lvlData.getRPr().getB()!=null && lvlData.getRPr().getB().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										}
									} else {
										if(lvlData.getRPr().getB()!=null && lvlData.getRPr().getB().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										}
									}
									
								} else {
									msg = new UserMessage();
									msg.setChkresultSuitable(false);
									msg.setFullySuitable(false);
									msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								}
								break;
							}
						}
					} else {
						msg = new UserMessage();
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	
	/*
	 * I
	 */
	public boolean rekursiveChkIInStyle(Style style, ArrayList<String> appearedStyleList){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getI()!=null){
					if(style.getRPr().getI().isVal()){
						result = true;
					} else {
						result = false;
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkIInStyle(basedOnStyle, appearedStyleList);
					}
					if(!result) {
						Style linkStyle = style.getLink()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkIInStyle(linkStyle, appearedStyleList);
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
	private boolean isStyleI(Style style){
		ArrayList<String> appearedStyleList = new ArrayList<String>();
		appearedStyleList.add(style.getStyleId());
		return rekursiveChkIInStyle(style, appearedStyleList);
	}
	public UserMessage isFontIAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		
		BigInteger numId = null;
		BigInteger lvl = null;
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal();
					}
					if(p.getPPr().getNumPr().getIlvl()!=null){
						lvl = p.getPPr().getNumPr().getIlvl().getVal();
					}
				} else {
					if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
						Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(p.getPPr().getPStyle().getVal());
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						numId = recursiveAbsNumLvl(style, stylePart, 0, appearedStyleList);
						appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						lvl = recursiveAbsNumLvl(style, stylePart, 1, appearedStyleList);
					}
				}
				if(numId!=null){
					AbstractNum absNum = getAbstNum(numId);
					if(absNum!=null){
						if(lvl==null){
							lvl = new BigInteger("1");
						}
						List<Lvl> lvlList = absNum.getLvl();
						for(int i=0; i<lvlList.size(); i++){
							if(lvlList.get(i)!=null && lvlList.get(i).getIlvl()!=null && lvlList.get(i).getIlvl().intValue() == lvl.intValue()){
								Lvl lvlData = lvlList.get(i);
								if(lvlData.getRPr()!=null){
									if(isStyleI(predefinedStyle)){
										if(lvlData.getRPr().getI()!=null && lvlData.getRPr().getI().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										}
									} else {
										if(lvlData.getRPr().getI()!=null && lvlData.getRPr().getI().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										}
									}
								} else {
									msg = new UserMessage();
									msg.setChkresultSuitable(false);
									msg.setFullySuitable(false);
									msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								}
								break;
							}
						}
					} else {
						msg = new UserMessage();
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	/*
	 * caps
	 */
	public boolean rekursiveChkCapsInStyle(Style style, ArrayList<String> appearedStyleList){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getCaps()!=null){
					if(style.getRPr().getCaps().isVal()){
						result = true;
					} else {
						result = false;
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkCapsInStyle(basedOnStyle, appearedStyleList);
					}
					if(!result) {
						Style linkStyle = style.getLink()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkCapsInStyle(linkStyle, appearedStyleList);
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
	private boolean isStyleCaps(Style style){
		ArrayList<String> appearedStyleList = new ArrayList<String>();
		appearedStyleList.add(style.getStyleId());
		return rekursiveChkCapsInStyle(style, appearedStyleList);
	}
	public UserMessage isFontCapsAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		
		BigInteger numId = null;
		BigInteger lvl = null;
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal();
					}
					if(p.getPPr().getNumPr().getIlvl()!=null){
						lvl = p.getPPr().getNumPr().getIlvl().getVal();
					}
				} else {
					if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
						Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(p.getPPr().getPStyle().getVal());
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						numId = recursiveAbsNumLvl(style, stylePart, 0, appearedStyleList);
						appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						lvl = recursiveAbsNumLvl(style, stylePart, 1, appearedStyleList);
					}
				}
				if(numId!=null){
					AbstractNum absNum = getAbstNum(numId);
					if(absNum!=null){
						if(lvl==null){
							lvl = new BigInteger("1");
						}
						List<Lvl> lvlList = absNum.getLvl();
						for(int i=0; i<lvlList.size(); i++){
							if(lvlList.get(i)!=null && lvlList.get(i).getIlvl()!=null && lvlList.get(i).getIlvl().intValue() == lvl.intValue()){
								Lvl lvlData = lvlList.get(i);
								if(lvlData.getRPr()!=null){
									if(isStyleCaps(predefinedStyle)){
										if(lvlData.getRPr().getCaps()!=null && lvlData.getRPr().getCaps().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										}
									} else {
										if(lvlData.getRPr().getCaps()!=null && lvlData.getRPr().getCaps().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										}
									}
									
								} else {
									msg = new UserMessage();
									msg.setChkresultSuitable(false);
									msg.setFullySuitable(false);
									msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								}
								break;
							}
						}
					} else {
						msg = new UserMessage();
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	/*
	 * smallCaps
	 */
	public boolean rekursiveChkSmallCapsInStyle(Style style, ArrayList<String> appearedStyleList){
		boolean result = false;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getSmallCaps()!=null){
					if(style.getRPr().getSmallCaps().isVal()){
						result = true;
					} else {
						result = false;
					}
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = false;
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkSmallCapsInStyle(basedOnStyle, appearedStyleList);
					}
					if(!result) {
						Style linkStyle = style.getLink()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = false;
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkSmallCapsInStyle(linkStyle, appearedStyleList);
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
	private boolean isStyleSmallCaps(Style style){
		ArrayList<String> appearedStyleList = new ArrayList<String>();
		appearedStyleList.add(style.getStyleId());
		return rekursiveChkSmallCapsInStyle(style, appearedStyleList);
	}
	public UserMessage isFontSmallCapsAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		
		BigInteger numId = null;
		BigInteger lvl = null;
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal();
					}
					if(p.getPPr().getNumPr().getIlvl()!=null){
						lvl = p.getPPr().getNumPr().getIlvl().getVal();
					}
				} else {
					if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
						Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(p.getPPr().getPStyle().getVal());
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						numId = recursiveAbsNumLvl(style, stylePart, 0, appearedStyleList);
						appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						lvl = recursiveAbsNumLvl(style, stylePart, 1, appearedStyleList);
					}
				}
				if(numId!=null){
					AbstractNum absNum = getAbstNum(numId);
					if(absNum!=null){
						if(lvl==null){
							lvl = new BigInteger("1");
						}
						List<Lvl> lvlList = absNum.getLvl();
						for(int i=0; i<lvlList.size(); i++){
							if(lvlList.get(i)!=null && lvlList.get(i).getIlvl()!=null && lvlList.get(i).getIlvl().intValue() == lvl.intValue()){
								Lvl lvlData = lvlList.get(i);
								if(lvlData.getRPr()!=null){
									if(isStyleSmallCaps(predefinedStyle)){
										if(lvlData.getRPr().getSmallCaps()!=null && lvlData.getRPr().getSmallCaps().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										}
									} else {
										if(lvlData.getRPr().getSmallCaps()!=null && lvlData.getRPr().getSmallCaps().isVal()){
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										}
									}
									
								} else {
									msg = new UserMessage();
									msg.setChkresultSuitable(false);
									msg.setFullySuitable(false);
									msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								}
								break;
							}
						}
					} else {
						msg = new UserMessage();
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	/*
	 * sz
	 */
	public BigInteger rekursiveChkSzInStyle(Style style, ArrayList<String> appearedStyleList){
		BigInteger result = null;
		try {
			if(style!=null && style.getRPr()!=null){
				if(style.getRPr().getSz()!=null){
					result = style.getRPr().getSz().getVal();
				} else {
					Style basedOnStyle = style.getBasedOn()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getBasedOn().getVal()) : null;
					if(basedOnStyle!=null){
						if(appearedStyleList.indexOf(basedOnStyle.getStyleId())>=0){
							result = null;
							return result;
						}
						appearedStyleList.add(basedOnStyle.getStyleId());
						result = rekursiveChkSzInStyle(basedOnStyle, appearedStyleList);
					}
					if(result==null) {
						Style linkStyle = style.getLink()!=null ? AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(style.getLink().getVal()) : null;
						if(linkStyle!=null){
							if(appearedStyleList.indexOf(linkStyle.getStyleId())>=0){
								result = null;
								return result;
							}
							appearedStyleList.add(linkStyle.getStyleId());
							result = rekursiveChkSzInStyle(linkStyle, appearedStyleList);
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
	private BigInteger getStyleSz(Style style){
		ArrayList<String> appearedStyleList = new ArrayList<String>();
		appearedStyleList.add(style.getStyleId());
		return rekursiveChkSzInStyle(style, appearedStyleList);
	}
	public UserMessage isFontSzAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		
		BigInteger numId = null;
		BigInteger lvl = null;
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getNumPr()!=null){
					if(p.getPPr().getNumPr().getNumId()!=null){
						numId = p.getPPr().getNumPr().getNumId().getVal();
					}
					if(p.getPPr().getNumPr().getIlvl()!=null){
						lvl = p.getPPr().getNumPr().getIlvl().getVal();
					}
				} else {
					if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
						Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(sourceFormat).getStyleMap().get(p.getPPr().getPStyle().getVal());
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						numId = recursiveAbsNumLvl(style, stylePart, 0, appearedStyleList);
						appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(p.getPPr().getPStyle().getVal());
						lvl = recursiveAbsNumLvl(style, stylePart, 1, appearedStyleList);
					}
				}
				if(numId!=null){
					AbstractNum absNum = getAbstNum(numId);
					if(absNum!=null){
						if(lvl==null){
							lvl = new BigInteger("1");
						}
						List<Lvl> lvlList = absNum.getLvl();
						for(int i=0; i<lvlList.size(); i++){
							if(lvlList.get(i)!=null && lvlList.get(i).getIlvl()!=null && lvlList.get(i).getIlvl().intValue() == lvl.intValue()){
								Lvl lvlData = lvlList.get(i);
								BigInteger chk1 = getStyleSz(predefinedStyle);
								if(chk1!=null){
									if(lvlData.getRPr().getSz()!=null){
										BigInteger chk2 = lvlData.getRPr().getSz().getVal();
										if(chk1.intValue() == chk2.intValue()){
											msg = new UserMessage();
											msg.setChkresultSuitable(true);
											msg.setFullySuitable(true);
											msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
										} else {
											msg = new UserMessage();
											msg.setChkresultSuitable(false);
											msg.setFullySuitable(false);
											msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
										}
									}
								} else {
									if(lvlData.getRPr()!=null && lvlData.getRPr().getSz()!=null && lvlData.getRPr().getSz().getVal()!=null){
										msg = new UserMessage();
										msg.setChkresultSuitable(false);
										msg.setFullySuitable(false);
										msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
									} else {
										msg = new UserMessage();
										msg.setChkresultSuitable(true);
										msg.setFullySuitable(true);
										msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
									}
								}
								break;
							}
						}
					} else {
						msg = new UserMessage();
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
}
