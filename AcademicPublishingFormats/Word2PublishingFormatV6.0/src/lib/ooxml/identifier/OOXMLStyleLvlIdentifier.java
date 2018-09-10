package lib.ooxml.identifier;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.CTSimpleField;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Style;
import org.docx4j.wml.DocDefaults.RPrDefault;

import base.UserMessage;
import tools.LogUtils;

public class OOXMLStyleLvlIdentifier extends OOXMLIdentifierProvider implements Identifier {

	private RPrDefault rPrDefault;
	
	public OOXMLStyleLvlIdentifier(String sourceFormat, RPrDefault rPrDefault) {
		super(sourceFormat);
		// TODO Auto-generated constructor stub
		this.rPrDefault = rPrDefault;
	}
	private UserMessage isBAccepted(Style rStyle, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(rStyle!=null){
				if(rStyle.getStyleId()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(rStyle.getStyleId());
					boolean chk1 = rekursiveChkBInStyle(rStyle, stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getB()!=null){
							chk1 = rPrDefault.getRPr().getB().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkBInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	public UserMessage isBAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(p.getPPr().getPStyle().getVal());
					boolean chk1 = rekursiveChkBInStyle(stylePart.getStyleById(p.getPPr().getPStyle().getVal()), stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getB()!=null){
							chk1 = rPrDefault.getRPr().getB().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkBInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 == chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	private UserMessage isIAccepted(Style rStyle, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(rStyle!=null){
				if(rStyle.getStyleId()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(rStyle.getStyleId());
					boolean chk1 = rekursiveChkIInStyle(rStyle, stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getI()!=null){
							chk1 = rPrDefault.getRPr().getI().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkIInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isIAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(p.getPPr().getPStyle().getVal());
					boolean chk1 = rekursiveChkIInStyle(stylePart.getStyleById(p.getPPr().getPStyle().getVal()), stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getI()!=null){
							chk1 = rPrDefault.getRPr().getI().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkIInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	private UserMessage isCapsAccepted(Style rStyle, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(rStyle!=null){
				if(rStyle.getStyleId()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(rStyle.getStyleId());
					boolean chk1 = rekursiveChkCapsInStyle(rStyle, stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getCaps()!=null){
							chk1 = rPrDefault.getRPr().getCaps().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkCapsInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isCapsAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(p.getPPr().getPStyle().getVal());
					boolean chk1 = rekursiveChkCapsInStyle(stylePart.getStyleById(p.getPPr().getPStyle().getVal()), stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getCaps()!=null){
							chk1 = rPrDefault.getRPr().getCaps().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkCapsInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	private UserMessage isSmallCapsAccepted(Style rStyle, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		msg = new UserMessage();
		try {
			if(rStyle!=null){
				if(rStyle.getStyleId()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(rStyle.getStyleId());
					boolean chk1 = rekursiveChkSmallCapsInStyle(rStyle, stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getSmallCaps()!=null){
							chk1 = rPrDefault.getRPr().getSmallCaps().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkSmallCapsInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isSmallCapsAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(p!=null){
				if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
					ArrayList<String> appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(p.getPPr().getPStyle().getVal());
					boolean chk1 = rekursiveChkSmallCapsInStyle(stylePart.getStyleById(p.getPPr().getPStyle().getVal()), stylePart, msg, appearedStylelist, true);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk1 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						if(rPrDefault==null){
							rPrDefault = super.getDocDefaultRPr(stylePart);
						}
						if(rPrDefault!=null 
								&& rPrDefault.getRPr()!=null 
								&& rPrDefault.getRPr().getSmallCaps()!=null){
							chk1 = rPrDefault.getRPr().getSmallCaps().isVal();
						} else {
							chk1 = false;
						}
					}
					UserMessage chkMsg1 = msg;
					
					msg = new UserMessage();
					appearedStylelist = new ArrayList<String>();
					appearedStylelist.add(predefinedStyle.getStyleId());
					boolean chk2 = rekursiveChkSmallCapsInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
					if(UserMessage.CHKRESULT_UNACCEPTED.equals(msg.getMessageCode())){
						chk2 = false;
					} else if(UserMessage.CHKRESULT_NULL.equals(msg.getMessageCode())){
						chk2 = false;
					}
					UserMessage chkMsg2 = msg;
					
					if(chk1 && chk2){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						if(chkMsg1!=null && chkMsg2!=null && UserMessage.CHKRESULT_NULL.equals(chkMsg1.getMessageCode()) && UserMessage.CHKRESULT_NULL.equals(chkMsg2.getMessageCode())){
							msg.setChkresultSuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	private UserMessage isFontSzAccepted(Style rStyle, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		msg = new UserMessage();
		try {
			if(rStyle!=null){
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(rStyle.getStyleId());
				BigDecimal chk1 = rekursiveGetSzInStyle(rStyle, stylePart, msg, appearedStylelist, true);
				if(chk1==null){
					if(rPrDefault==null){
						rPrDefault = super.getDocDefaultRPr(stylePart);
					}
					if(rPrDefault!=null 
							&& rPrDefault.getRPr()!=null 
							&& rPrDefault.getRPr().getSz()!=null){
						chk1 = BigDecimal.valueOf(rPrDefault.getRPr().getSz().getVal().intValue());
					}
				}
				msg = new UserMessage();
				appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				BigDecimal chk2 = rekursiveGetSzInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
				
				if(chk1!=null && chk2!=null){
					if(chk1.intValue() == chk2.intValue()){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						msg.setChkresultSuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
					}
				} else {
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
				}
			} else {
				msg.setChkresultSuitable(false);
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isFontSzAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		msg = new UserMessage();
		try {
			if(p.getPPr()!=null && p.getPPr().getPStyle()!=null){
				ArrayList<String> appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(p.getPPr().getPStyle().getVal());
				BigDecimal chk1 = rekursiveGetSzInStyle(stylePart.getStyleById(p.getPPr().getPStyle().getVal()), stylePart, msg, appearedStylelist, true);
				if(chk1 == null){
//					PPrDefault pPrDefault = super.getDocDefaultPPr(stylePart);
					if(rPrDefault==null){
						rPrDefault = super.getDocDefaultRPr(stylePart);
					}
					if(rPrDefault!=null 
							&& rPrDefault.getRPr()!=null 
							&& rPrDefault.getRPr().getSz()!=null){
						chk1 = BigDecimal.valueOf(rPrDefault.getRPr().getSz().getVal().intValue());
					}
					if(chk1==null){
						chk1 = new BigDecimal("20");
					}
				}
				msg = new UserMessage();
				appearedStylelist = new ArrayList<String>();
				appearedStylelist.add(predefinedStyle.getStyleId());
				BigDecimal chk2 = rekursiveGetSzInStyle(predefinedStyle, stylePart, msg, appearedStylelist, false);
				
				if(chk1!=null && chk2!=null){
					if(chk1.intValue() == chk2.intValue()){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					} else {
						msg.setChkresultSuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
					}
				} else {
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
				}
			} else {
				msg.setChkresultSuitable(false);
				msg.setMessageCode(UserMessage.CHKRESULT_NULL);
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	
	/*
	 * r style
	 */
	public UserMessage isRBAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, boolean isDeep){
		UserMessage msg = new UserMessage();
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			if(p!=null){
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
							Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
							if(isBAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
								if(msg.getBeginPosition()==-1){
									msg.setBeginPosition(i);
								}
								msg.setEndPosition(i);
							} else {
								break;
							}
						}
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R r = (R)sfO;
										if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
											Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
											if(isBAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
												if(msg.getBeginPosition()==-1){
													msg.setBeginPosition(i);
												}
												msg.setEndPosition(i);
											} else {
												break;
											}
										}
									}
								}
							}
						}
					}
				}
				
				if(msg.getBeginPosition()==-1){
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_NULL);
				} else if(isDeep && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					if(msg.getEndPosition()+1==msg.getIndexOfLastR()){
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isRIAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, boolean isDeep){
		UserMessage msg = new UserMessage();
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			if(p!=null){
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
							Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
							if(isIAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
								if(msg.getBeginPosition()==-1){
									msg.setBeginPosition(i);
								}
								msg.setEndPosition(i);
							} else {
								break;
							}
						}
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R r = (R)sfO;
										if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
											Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
											if(isIAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
												if(msg.getBeginPosition()==-1){
													msg.setBeginPosition(i);
												}
												msg.setEndPosition(i);
											} else {
												break;
											}
										}
									}
								}
							}
						}
					}
				}
				
				if(msg.getBeginPosition()==-1){
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_NULL);
				} else if(isDeep && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					if(msg.getEndPosition()+1==msg.getIndexOfLastR()){
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isRCapsAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, boolean isDeep){
		UserMessage msg = new UserMessage();
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			if(p!=null){
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
							Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
							if(isCapsAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
								if(msg.getBeginPosition()==-1){
									msg.setBeginPosition(i);
								}
								msg.setEndPosition(i);
							} else {
								break;
							}
						}
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R r = (R)sfO;
										if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
											Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
											if(isCapsAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
												if(msg.getBeginPosition()==-1){
													msg.setBeginPosition(i);
												}
												msg.setEndPosition(i);
											} else {
												break;
											}
										}
									}
								}
							}
						}
					}
				}
				
				if(msg.getBeginPosition()==-1){
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_NULL);
				} else if(isDeep && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					if(msg.getEndPosition()+1==msg.getIndexOfLastR()){
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isRSmallCapsAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, boolean isDeep){
		UserMessage msg = new UserMessage();
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			if(p!=null){
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
							Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
							if(isSmallCapsAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
								if(msg.getBeginPosition()==-1){
									msg.setBeginPosition(i);
								}
								msg.setEndPosition(i);
							} else {
								break;
							}
						}
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R r = (R)sfO;
										if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
											Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
											if(isSmallCapsAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
												if(msg.getBeginPosition()==-1){
													msg.setBeginPosition(i);
												}
												msg.setEndPosition(i);
											} else {
												break;
											}
										}
									}
								}
							}
						}
					}
				}
				
				if(msg.getBeginPosition()==-1){
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_NULL);
				} else if(isDeep && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					if(msg.getEndPosition()+1==msg.getIndexOfLastR()){
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		return msg;
	}
	public UserMessage isRFontSzAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, boolean isDeep){
		UserMessage msg = new UserMessage();
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			if(p!=null){
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				for (int i = 0; i < p.getContent().size(); i++) {
					Object o = p.getContent().get(i);
					if(o instanceof R){
						R r = (R)o;
						if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
							Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
							if(isFontSzAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
								if(msg.getBeginPosition()==-1){
									msg.setBeginPosition(i);
								}
								msg.setEndPosition(i);
							} else {
								break;
							}
						}
					} else if(o instanceof JAXBElement<?>){
						if(((JAXBElement<?>)o).getValue() instanceof CTSimpleField){
							CTSimpleField sf = (CTSimpleField) ((JAXBElement<?>)o).getValue();
							List<Object> sfContent = sf.getContent();
							if(sfContent!=null){
								for (int j = 0; j < sfContent.size(); j++) {
									Object sfO = sfContent.get(j);
									if(sfO instanceof R){
										R r = (R)sfO;
										if(r.getRPr()!=null && r.getRPr().getRStyle()!=null){
											Style rStyle = stylePart.getStyleById(r.getRPr().getRStyle().getVal());
											if(isFontSzAccepted(rStyle, stylePart, predefinedStyle, msg).isChkresultSuitable()){
												if(msg.getBeginPosition()==-1){
													msg.setBeginPosition(i);
												}
												msg.setEndPosition(i);
											} else {
												break;
											}
										}
									}
								}
							}
						}
					}
				}
				
				if(msg.getBeginPosition()==-1){
					msg.setChkresultSuitable(false);
					msg.setMessageCode(UserMessage.CHKRESULT_NULL);
				} else if(isDeep && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
					if(msg.getEndPosition()+1==msg.getIndexOfLastR()){
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
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
