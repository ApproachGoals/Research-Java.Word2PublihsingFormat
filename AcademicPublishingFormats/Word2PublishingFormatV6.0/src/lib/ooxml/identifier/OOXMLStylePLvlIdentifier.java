package lib.ooxml.identifier;

import java.math.BigDecimal;
import java.util.ArrayList;

import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.P;
import org.docx4j.wml.Style;

import base.UserMessage;
import tools.LogUtils;

public class OOXMLStylePLvlIdentifier extends OOXMLIdentifierProvider implements Identifier {
	public OOXMLStylePLvlIdentifier(String sourceFormat) {
		super(sourceFormat);
		// TODO Auto-generated constructor stub
	}
	public UserMessage isBAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		try {
			if(p!=null){
				boolean chk1 = rekursiveChkBInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(p.getPPr()!=null && p.getPPr().getRPr()!=null){
					if(p.getPPr().getRPr().getB()!=null){
						if(p.getPPr().getRPr().getB().isVal()){
							if(chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						} else {
							if(!chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						}
					} else {
						if(!chk1){
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						}
					}
				} else {
					if(!chk1){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					} else {
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
	public UserMessage isIAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStylelist, boolean isFromDoc){
		try {
			if(p!=null){
				boolean chk1 = rekursiveChkIInStyle(chkingStyle, stylePart, msg, appearedStylelist, isFromDoc);
				if(p.getPPr()!=null && p.getPPr().getRPr()!=null){
					if(p.getPPr().getRPr().getI()!=null){
						if(p.getPPr().getRPr().getI().isVal()){
							if(chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						} else {
							if(!chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						}
					} else {
						if(!chk1){
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						}
					}
				} else {
					if(!chk1){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					} else {
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
	public UserMessage isCapsAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		try {
			if(p!=null){
				boolean chk1 = rekursiveChkCapsInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(p.getPPr()!=null && p.getPPr().getRPr()!=null){
					if(p.getPPr().getRPr().getCaps()!=null){
						if(p.getPPr().getRPr().getCaps().isVal()){
							if(chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						} else {
							if(!chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						}
					} else {
						if(!chk1){
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						}
					}
				} else {
					if(!chk1){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					} else {
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
	public UserMessage isSmallCapsAccepted(P p, StyleDefinitionsPart stylePart, Style chkingStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		try {
			if(p!=null){
				boolean chk1 = rekursiveChkSmallCapsInStyle(chkingStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(p.getPPr()!=null && p.getPPr().getRPr()!=null){
					if(p.getPPr().getRPr().getSmallCaps()!=null){
						if(p.getPPr().getRPr().getSmallCaps().isVal()){
							if(chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						} else {
							if(!chk1){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						}
					} else {
						if(!chk1){
							msg.setChkresultSuitable(true);
							msg.setFullySuitable(true);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						}
					}
				} else {
					if(!chk1){
						msg.setChkresultSuitable(true);
						msg.setFullySuitable(true);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					} else {
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
	public UserMessage isFontSzAccepted(P p, StyleDefinitionsPart stylePart, Style predefinedStyle, UserMessage msg, ArrayList<String> appearedStyleList, boolean isFromDoc){
		try {
			if(p!=null){
				BigDecimal chk1 = rekursiveGetSzInStyle(predefinedStyle, stylePart, msg, appearedStyleList, isFromDoc);
				if(p.getPPr()!=null && p.getPPr().getRPr()!=null){
					if(p.getPPr().getRPr().getSz()!=null && p.getPPr().getRPr().getSz().getVal()!=null){
						if(chk1!=null){
							if(chk1.intValue() == p.getPPr().getRPr().getSz().getVal().intValue()){
								msg.setChkresultSuitable(true);
								msg.setFullySuitable(true);
								msg.setMessageCode(UserMessage.CHKRESULT_ACCEPTED);
							} else {
								msg.setChkresultSuitable(false);
								msg.setFullySuitable(false);
								msg.setMessageCode(UserMessage.CHKRESULT_UNACCEPTED);
							}
						} else {
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						}
					} else {
						if(p.getPPr().getRPr().getSz()==null && chk1==null){
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						} else {
							msg.setChkresultSuitable(false);
							msg.setFullySuitable(false);
							msg.setMessageCode(UserMessage.CHKRESULT_NULL);
						}
					}
				} else {
					if(chk1!=null){
						msg.setChkresultSuitable(false);
						msg.setFullySuitable(false);
						msg.setMessageCode(UserMessage.CHKRESULT_NULL);
					} else {
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
