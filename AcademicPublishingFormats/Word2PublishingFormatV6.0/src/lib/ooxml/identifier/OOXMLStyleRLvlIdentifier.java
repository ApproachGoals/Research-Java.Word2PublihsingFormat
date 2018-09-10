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

import base.UserMessage;
import tools.LogUtils;

public class OOXMLStyleRLvlIdentifier extends OOXMLIdentifierProvider implements Identifier {
	
	public OOXMLStyleRLvlIdentifier(String sourceFormat) {
		super(sourceFormat);
		// TODO Auto-generated constructor stub
	}
	public static int CHK_LEFT_2_RIGHT_PART = 1;
	public static int CHK_LEFT_2_RIGHT_FULL = 2;
	
	public UserMessage isBAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		try {
			if(p!=null){
				msg = getAcceptedBTagIndex(p, stylePart, chkingStyle, msg);
				if(msg.getIndexOfLastR()==msg.getEndPosition()){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(true);
				} else if(msg.getIndexOfLastR()>msg.getEndPosition() && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(false);
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	public UserMessage isIAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		try {
			if(p!=null){
				msg = getAcceptedITagIndex(p, stylePart, chkingStyle, chkingStyle, msg);
				if(msg.getIndexOfLastR()==msg.getEndPosition()){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(true);
				} else if(msg.getIndexOfLastR()>msg.getEndPosition() && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(false);
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}

		return msg;
	}
	public UserMessage isCapsAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		try {
			if(p!=null){
				msg = getAcceptedCapsTagIndex(p, stylePart, chkingStyle, msg);
				if(msg.getIndexOfLastR()==msg.getEndPosition()){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(true);
				} else if(msg.getIndexOfLastR()>msg.getEndPosition() && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(false);
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}

		return msg;
	}
	public UserMessage isSmallCapsAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		try {
			if(p!=null){
				msg = getAcceptedSmallCapsTagIndex(p, stylePart, chkingStyle, msg);
				if(msg.getIndexOfLastR()==msg.getEndPosition() && UserMessage.CHKRESULT_ACCEPTED.equals(msg.getMessageCode())){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(true);
				} else if(msg.getIndexOfLastR()>msg.getEndPosition() && msg.getBeginPosition()==0){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(false);
				}
			}
			
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}

		
		return msg;
	}
	public UserMessage isFontSzAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		try {
			if(p!=null){
				msg = getAcceptedFontSzTagIndex(p, stylePart, predefinedStyle, msg);
				if(msg.getIndexOfLastR()==msg.getEndPosition() && UserMessage.CHKRESULT_ACCEPTED.equals(msg.getMessageCode())){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(true);
				} else if(msg.getIndexOfLastR()>msg.getEndPosition() && msg.getBeginPosition()==0 && UserMessage.CHKRESULT_ACCEPTED.equals(msg.getMessageCode())){
					msg.setChkresultSuitable(true);
					msg.setFullySuitable(false);
				} else {
					msg.setChkresultSuitable(false);
					msg.setFullySuitable(false);
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}

		return msg;
	}
	
	/*
	 * get position
	 */
	public UserMessage getAcceptedBTagIndex(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						R r = (R)o;
						ArrayList<String> appearedStylelist = new ArrayList<String>();
						appearedStylelist.add(chkingStyle.getStyleId());
						if(isBAccepted(r, stylePart, chkingStyle, msg, appearedStylelist, false)){
							if(msg.getBeginPosition()==-1){
								msg.setBeginPosition(i);
							}
							msg.setEndPosition(i);
						} else {
							break;
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
										ArrayList<String> appearedStylelist = new ArrayList<String>();
										appearedStylelist.add(chkingStyle.getStyleId());
										if(isBAccepted(sfR, stylePart, chkingStyle, msg, appearedStylelist, false)){
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
					} else {
//						System.out.println(o.getClass().getName());
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	public UserMessage getAcceptedITagIndex(P p, StyleDefinitionsPart stylePart, Style chkingStyle, Style predefinedStyle, UserMessage msg){
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						R r = (R)o;
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(predefinedStyle.getStyleId());
						if(isIAccepted(r, stylePart, predefinedStyle, msg, appearedStyleList, false)){
							if(msg.getBeginPosition()==-1){
								msg.setBeginPosition(i);
							}
							msg.setEndPosition(i);
						} else {
							break;
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
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(predefinedStyle.getStyleId());
										if(isIAccepted(sfR, stylePart, predefinedStyle, msg, appearedStyleList, false)){
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
					} else {

					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		
		return msg;
	}
	public UserMessage getAcceptedCapsTagIndex(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						R r = (R)o;
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(chkingStyle.getStyleId());
						if(isCapsAccepted(r, stylePart, chkingStyle, msg, appearedStyleList, false)){
							if(msg.getBeginPosition()==-1){
								msg.setBeginPosition(i);
							}
							msg.setEndPosition(i);
						} else {
							break;
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
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(chkingStyle.getStyleId());
										if(isCapsAccepted(sfR, stylePart, chkingStyle, msg, appearedStyleList, false)){
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
					} else {

					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		
		return msg;
	}
	public UserMessage getAcceptedSmallCapsTagIndex(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg){
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						R r = (R)o;
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(chkingStyle.getStyleId());
						if(isSmallCapsAccepted(r, stylePart, chkingStyle, msg, appearedStyleList, false)){
							if(msg.getBeginPosition()==-1){
								msg.setBeginPosition(i);
							}
							msg.setEndPosition(i);
						} else {
							break;
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
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(chkingStyle.getStyleId());
										if(isSmallCapsAccepted(sfR, stylePart, chkingStyle, msg, appearedStyleList, false)){
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
					} else {

					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		
		
		return msg;
	}
	public UserMessage getAcceptedFontSzTagIndex(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg){
		msg.setBeginPosition(-1);
		msg.setEndPosition(-1);
		msg.setIndexOfLastR(-1);
		try {
			List<Object> pContents = p.getContent();
			if(pContents!=null && pContents.size()>0){
				for (int i = 0; i < pContents.size(); i++) {
					if(pContents.get(i) instanceof R){
						msg.setIndexOfLastR(i);
					}
				}
				
				for (int i = 0; i < pContents.size(); i++) {
					Object o = pContents.get(i);
					if(o instanceof R){
						R r = (R)o;
						ArrayList<String> appearedStyleList = new ArrayList<String>();
						appearedStyleList.add(predefinedStyle.getStyleId());
						if(isFontSzAccepted(r, stylePart, predefinedStyle, msg, appearedStyleList, false)){
							if(msg.getBeginPosition()==-1){
								msg.setBeginPosition(i);
							}
							msg.setEndPosition(i);
						} else {
							break;
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
										ArrayList<String> appearedStyleList = new ArrayList<String>();
										appearedStyleList.add(predefinedStyle.getStyleId());
										if(isFontSzAccepted(sfR, stylePart, predefinedStyle, msg, appearedStyleList, false)){
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
					} else {

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
	 * check single r block
	 */
	public boolean isBAccepted(R r, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(r!=null){
				boolean chk1 = rekursiveChkBInStyle(chkingStyle, stylePart, msg, appearedStyleList, false);
				if(r.getRPr()!=null){
					if(r.getRPr().getB()!=null){
						if(r.getRPr().getB().isVal()){
							if(chk1){
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
							if(!chk1){
								result = true;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(false);
							} else {
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
							}
						}
					} else {
						if(rekursiveChkStyle(chkingStyle, stylePart, true, false, false, false)){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else if(!chk1){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else {
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						}
					}
				} else {
					if(!chk1){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(false);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return result;
	}
	public boolean isIAccepted(R r, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(r!=null){
				boolean chk1 = rekursiveChkIInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(r.getRPr()!=null){
					if(r.getRPr().getI()!=null){
						if(r.getRPr().getI().isVal()){
							if(chk1){
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
							if(!chk1){
								result = true;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(false);
							} else {
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
							}
						}
					} else {
						if(rekursiveChkStyle(chkingStyle, stylePart, false, true, false, false)){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else if(!chk1){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else {
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						}
					}
				} else {
					if(!chk1){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(false);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return result;
	}
	public boolean isCapsAccepted(R r, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(r!=null){
				boolean chk1 = rekursiveChkCapsInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(r.getRPr()!=null){
					if(r.getRPr().getCaps()!=null){
						if(r.getRPr().getCaps().isVal()){
							if(chk1){
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
							if(!chk1){
								result = true;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(false);
							} else {
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
							}
						}
					} else {
						if(rekursiveChkStyle(chkingStyle, stylePart, false, false, true, false)){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else if(!chk1){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else {
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						}
					}
				} else {
					if(!chk1){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(false);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return result;
	}
	public boolean isSmallCapsAccepted(R r, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(r!=null){
				boolean chk1 = rekursiveChkSmallCapsInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(r.getRPr()!=null){
					if(r.getRPr().getSmallCaps()!=null){
						if(r.getRPr().getSmallCaps().isVal()){
							if(chk1){
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
							if(!chk1){
								result = true;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(false);
							} else {
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
							}
						}
					} else {
						if(rekursiveChkStyle(chkingStyle, stylePart, false, false, false, true)){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else if(!chk1){
							result = true;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(false);
						} else {
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						}
					}
				} else {
					if(!chk1){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(false);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return result;
	}
	public boolean isFontSzAccepted(R r, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		boolean result = false;
		try {
			if(r!=null){
				BigDecimal chk1 = rekursiveGetSzInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(r.getRPr()!=null){
					if(r.getRPr().getSz()!=null && r.getRPr().getSz().getVal()!=null){
						if(chk1!=null){
							if(chk1.intValue() == r.getRPr().getSz().getVal().intValue()){
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
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						}
					} else if(r.getRPr().getSzCs()!=null && r.getRPr().getSzCs().getVal()!=null) {
						if(chk1.intValue() == r.getRPr().getSzCs().getVal().intValue()){
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
						if(chk1!=null){
							if(r.getRPr().getSz()==null && chk1==null){
								result = true;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
							} else {
								result = false;
								msg.setMessageCode(UserMessage.CHKRESULT_NULL);
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
							}
						} else {
							result = false;
							msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
						}
					}
				} else {
					if(chk1!=null){
						result = true;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(false);
					} else {
						result = false;
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return result;
	}
}
