package lib.ooxml.converter;

import java.io.File;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FootnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTFootnotes;
import org.docx4j.wml.CTFtnEdn;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;
import org.docx4j.wml.DocDefaults.RPrDefault;
import org.docx4j.wml.PPrBase.PStyle;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import lib.ooxml.author.Authors;
import lib.ooxml.tool.OOXMLConvertTool;
import tools.LogUtils;

public class OOXMLConverterSpringer implements OOXMLConverter{

	private Map<String, Style> springerStyle;
	
	public OOXMLConverterSpringer() {
		// TODO Auto-generated constructor stub
		springerStyle = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap();
	}
	
	@Override
	public UserMessage adaptAbstract(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(p!=null){
				String pText = OOXMLConvertTool.getPText(p);
				if(pText!=null && pText.replaceAll(" ", "").length()>0){
					if(OOXMLConvertTool.isPTextStartWithString(p, "Abstract", false) && OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "").equals("abstract")) {
						Text text = Context.getWmlObjectFactory().createText();
						text.setValue("Abstract.");
						R newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						newR.setRPr(Context.getWmlObjectFactory().createRPr());
						newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.ABSTRACTHEADER).getStyleId());
						
						P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
						abstractTextP.setPPr(Context.getWmlObjectFactory().createPPr());
						OOXMLConvertTool.removeAllRPROfP(abstractTextP);
						abstractTextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						abstractTextP.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.ABSTRACT).getStyleId());
						abstractTextP.getContent().add(0, newR);
						
						documentPart.getContent().remove(docContentIndex);
					} else if(OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)8212), true)
							&& OOXMLConvertTool.isPTextStartWithString(p, "Abstract"+((char)8212), false)){
						pText = pText.replaceFirst("Abstract"+((char)8212), "");
						Text text = Context.getWmlObjectFactory().createText();
						text.setValue("Abstract.");
						R newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						newR.setRPr(Context.getWmlObjectFactory().createRPr());
						newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.ABSTRACTHEADER).getStyleId());
						newR.getRPr().setB(new BooleanDefaultTrue());
						OOXMLConvertTool.removeStringFromPLostRStyle(p, "Abstract"+((char)8212), false);
						if(p.getPPr()!=null){
							
						} else {
							p.setPPr(Context.getWmlObjectFactory().createPPr());
						}
						p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.ABSTRACT).getStyleId());
						p.getContent().add(0, newR);
					}
				}
				msg.setMessageCode(UserMessage.MESSAGE_ABSTRACT_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptAcknowledgment(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(p!=null){
				String pText = OOXMLConvertTool.getPText(p);
				if(pText!=null && pText.replaceAll(" ", "").length()>0){
					if(OOXMLConvertTool.isPTextStartWithString(p, "Acknowledgments", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Acknowledgments")	// Elsevier
							|| OOXMLConvertTool.isPTextStartWithString(p, "Acknowledgment", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Acknowledgment")	// IEEE
							|| OOXMLConvertTool.isPTextStartWithString(p, "ACKNOWLEDGMENTS", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("ACKNOWLEDGMENTS")	// ACM
							) {
						Text text = Context.getWmlObjectFactory().createText();
						text.setValue("Acknowledgments.");
						R newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						newR.setRPr(Context.getWmlObjectFactory().createRPr());
						newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER).getStyleId());
						newR.getRPr().setB(new BooleanDefaultTrue());
						
						P aknowledgmentText = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
						
						aknowledgmentText.getContent().add(0, newR);
						
						documentPart.getContent().remove(docContentIndex);
						
					}
					
					msg.setMessageCode(UserMessage.MESSAGE_ACKNOWLEDGMENT_ADAPT_FINISH);
					if(AppEnvironment.getInstance().isActiveTestMode()){
						LogUtils.log(msg.getMessage());
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptAppendixHeader(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptAuthors(MainDocumentPart documentPart, int docContentIndex, int authorCount,
			Authors authors) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptFigureCaption(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			if(p!=null){
				if(p.getPPr()!=null){
					
				} else {
					p.setPPr(Context.getWmlObjectFactory().createPPr());
				}
				if(p.getPPr().getPStyle()!=null){
					
				} else {
					p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				}
				p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.FIGURE).getStyleId());
//				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
//				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
//				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				OOXMLConvertTool.removeFirstStr(p, "Fig. ", ".");
				OOXMLConvertTool.removeFirstStr(p, "Figure ", ".");
//				OOXMLConvertTool.adaptFirstWord2Special(p, "Figure", "Fig.", springerStyle.get(AcademicFormatStructureDefinition.FIGURE).getStyleId());
				
				msg.setMessageCode(UserMessage.MESSAGE_CAPTION_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage adaptFontStyleClass(File convertedFile, WordprocessingMLPackage wordMLPackage,
			MainDocumentPart documentPart, StyleDefinitionsPart stylePart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try{
			for(Style o : stylePart.getContents().getStyle()){
				if(o.isDefault() && o.getType()!=null && "paragraph".equals(o.getType())){
					o.setDefault(false);
					break;
				}
			}
		} catch (Exception e){
			LogUtils.log("Error while changing the default style status in style class file.");
			LogUtils.log(e.getMessage());
		}
		
		Iterator<Map.Entry<String, Style>> entries = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap().entrySet().iterator();
		
		while (entries.hasNext()) {
			Map.Entry<String, Style> entry = entries.next();
			Style styleEntry = entry.getValue();
			try {
				if(OOXMLConvertTool.isStyleExist(styleEntry, stylePart)){
					Style style = stylePart.getStyleById(styleEntry.getStyleId());
					int styleIndex = style!=null?OOXMLConvertTool.getStyleDefinitionIndex(style, stylePart):-1;
					if(styleIndex>=0){
						stylePart.getContents().getStyle().remove(styleIndex);
						stylePart.getContents().getStyle().add(styleIndex, styleEntry);
					}
				} else {
					stylePart.getContents().getStyle().add(styleEntry);
				}
				
				
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log("Error while looping styles of style class file.");
				LogUtils.log(e.getMessage());
			}
		}
		try {
			List<Object> jaxbNodes = stylePart.getJAXBNodesViaXPath("w:docDefaults", false);
			if(jaxbNodes!=null){
				for (int i = 0; i < jaxbNodes.size(); i++) {
					if(jaxbNodes.get(i) instanceof DocDefaults){
						RPrDefault rPrDefault = (((DocDefaults) jaxbNodes.get(i)).getRPrDefault());
						Style defaultStyle = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap().get(AcademicFormatStructureDefinition.RPRDEFAULT);
						if(defaultStyle!=null && defaultStyle.getRPr()!=null){
							rPrDefault.setRPr(defaultStyle.getRPr());
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log("Error while changing the docDefault settings in style class file.");
			LogUtils.log(e.getMessage());
		}
		msg.setMessageCode(UserMessage.MESSAGE_STYLECLASS_ADAPT_FINISH);
		if(AppEnvironment.getInstance().isActiveTestMode()){
			LogUtils.log(msg.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage adaptFootNote(File convertedFile, WordprocessingMLPackage wordMLPackage) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			FootnotesPart footPart = wordMLPackage.getMainDocumentPart().getFootnotesPart();
			if(footPart!=null){
				CTFootnotes footNotes = footPart.getContents();
				if(footNotes!=null){
					List<CTFtnEdn> footNotesList =  footNotes.getFootnote();
					for (int i = 0; i < footNotesList.size(); i++) {
						CTFtnEdn footNote = footNotesList.get(i);
						List<Object> footNoteContent = footNote.getContent();
						if(footNoteContent!=null){
							for (Object o : footNoteContent) {
								if(o instanceof P){
									P p = (P)o;
									if(p.getPPr()!=null){
										p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
										p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.FOOTNOTE).getStyleId());
										
										if(p.getPPr().getRPr()!=null){
											p.getPPr().getRPr().setB(null);
											p.getPPr().getRPr().setBCs(null);
											p.getPPr().getRPr().setI(null);
											p.getPPr().getRPr().setICs(null);
											p.getPPr().getRPr().setCaps(null);
											p.getPPr().getRPr().setSmallCaps(null);
											p.getPPr().getRPr().setRFonts(null);
										}
										
									}
								}
							}
						}
					}
				}
			}
			msg.setMessageCode(UserMessage.MESSAGE_FOOTNOTE_ADAPT_FINISH);
			if(AppEnvironment.getInstance().isActiveTestMode()){
				LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading1(P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		if(p!=null){
			try {
				OOXMLConvertTool.removeAllProperties(p);
				OOXMLConvertTool.removeSectionNumStringOfHeading(p);
				OOXMLConvertTool.convertUppcaseToNormalWord(p);
				p.setPPr(Context.getWmlObjectFactory().createPPr());
				p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING1).getStyleId());
				
				msg.setMessageCode(UserMessage.MESSAGE_HEADING1_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading2(P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		if(p!=null){
			try {
				PPr ppr = OOXMLConvertTool.getPPr(p, true, true);
				if(ppr.getPStyle()!=null){
					
				} else {
					ppr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				}
				ppr.getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING2).getStyleId());
				OOXMLConvertTool.removeSectionNumStringOfHeading(p);
				
				msg.setMessageCode(UserMessage.MESSAGE_HEADING2_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading3(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		if(p!=null){
			try {
				if(OOXMLConvertTool.isAllRInSameStyle(p, documentPart.getStyleDefinitionsPart())){
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
					
					String pText = OOXMLConvertTool.getPText(p);
					P nextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
					R newR = Context.getWmlObjectFactory().createR();
					newR.getContent().add(Context.getWmlObjectFactory().createRLastRenderedPageBreak());
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue(pText+". ");
					text.setSpace("preserve");
					newR.getContent().add(text);
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING3CHAR).getStyleId());
					
					if(nextP.getPPr()!=null){
						
					} else {
						nextP.setPPr(Context.getWmlObjectFactory().createPPr());
					}
//					nextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
//					nextP.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING3).getStyleId());
					nextP.getPPr().setPStyle(null);
//					nextP.getPPr().setRPr(null);
					
					nextP.getContent().add(0, newR);
					
					documentPart.getContent().remove(docContentIndex);
				} else {
					
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
//					Style rStyle = OOXMLConvertTool.getFirstRStyle(p);
//					OOXMLConvertTool.replaceFirstSameGroupRStyle(p, rStyle, springerStyle.get(AcademicFormatStructureDefinition.HEADING3CHAR), ":", ".");
					OOXMLConvertTool.replaceFirstSameGroupRStyle(p, springerStyle.get(AcademicFormatStructureDefinition.HEADING3CHAR), documentPart.getStyleDefinitionsPart(), ":", ".");
					
					if(p.getPPr()!=null){
						
					} else {
						p.setPPr(Context.getWmlObjectFactory().createPPr());
					}
//					p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
//					p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING3).getStyleId());
					p.getPPr().setPStyle(null);
//					p.getPPr().setRPr(null);
				}
				msg.setMessageCode(UserMessage.MESSAGE_HEADING3_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading4(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		if(p!=null){
			try {
				if(OOXMLConvertTool.isAllRInSameStyle(p, documentPart.getStyleDefinitionsPart())){
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
					
					String pText = OOXMLConvertTool.getPText(p);
					P nextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue(pText+". ");
					text.setSpace("preserve");
					newR.getContent().add(text);
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING4CHAR).getStyleId());
					
					if(nextP.getPPr()!=null){
						
					} else {
						nextP.setPPr(Context.getWmlObjectFactory().createPPr());
					}
					nextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
//					nextP.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING4).getStyleId());
					nextP.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.NORMAL).getStyleId());
					nextP.getPPr().setRPr(null);
					
					nextP.getContent().add(0, newR);
					
					documentPart.getContent().remove(docContentIndex);
				} else {
					
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
//					Style rStyle = OOXMLConvertTool.getFirstRStyle(p);
//					OOXMLConvertTool.replaceFirstSameGroupRStyle(p, rStyle, springerStyle.get(AcademicFormatStructureDefinition.HEADING4CHAR), ":", ".");
					OOXMLConvertTool.replaceFirstSameGroupRStyle(p, springerStyle.get(AcademicFormatStructureDefinition.HEADING4CHAR), documentPart.getStyleDefinitionsPart(), ":", ".");
					if(p.getPPr()!=null){
						
					} else {
						p.setPPr(Context.getWmlObjectFactory().createPPr());
					}
					p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
//					p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.HEADING4).getStyleId());
					p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.NORMAL).getStyleId());
				}
				msg.setMessageCode(UserMessage.MESSAGE_HEADING4_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptImages(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptKeyWord(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(p!=null){
				String pText = OOXMLConvertTool.getPText(p);
				if(pText!=null && pText.replaceAll(" ", "").length()>0){
					if(OOXMLConvertTool.isPTextStartWithString(p, "Keywords", false) && OOXMLConvertTool.getPText(p).toLowerCase().replaceAll(" ", "").equals("keywords")) {
						Text text = Context.getWmlObjectFactory().createText();
						text.setValue("Keywords:");
						R newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						newR.setRPr(Context.getWmlObjectFactory().createRPr());
						newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.KEYWORDHEADER).getStyleId());
						
						P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
						abstractTextP.setPPr(Context.getWmlObjectFactory().createPPr());
						OOXMLConvertTool.removeAllRPROfP(abstractTextP);
						abstractTextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						abstractTextP.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.KEYWORDS).getStyleId());
						abstractTextP.getContent().add(0, newR);
						
						documentPart.getContent().remove(docContentIndex);
					} else if(OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)8212), true)
							||OOXMLConvertTool.isPTextIncludeString(p, "Keywords-", true)){
//						pText = pText.replaceFirst("Keywords"+((char)8212), "");
						pText = pText.substring(pText.indexOf("Keywords")+"Keywords".length()+1, pText.length());
						p.getContent().clear();
						Text text = Context.getWmlObjectFactory().createText();
						text.setValue("Keywords:");
						R newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						newR.setRPr(Context.getWmlObjectFactory().createRPr());
						newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.KEYWORDHEADER).getStyleId());
						
//						OOXMLConvertTool.removeStringFromPLostRStyle(p, "Keywords:");
//						OOXMLConvertTool.removeAllRPROfP(p);
						p.getContent().add(newR);
						
						text = Context.getWmlObjectFactory().createText();
						text.setValue(pText);
						newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						
						p.getContent().add(newR);
						
						if(p.getPPr()!=null){
							
						} else {
							p.setPPr(Context.getWmlObjectFactory().createPPr());
						}
						p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.KEYWORDS).getStyleId());
//						p.getContent().add(0, newR);
					} else if(OOXMLConvertTool.isPTextIncludeString(p, "Keywords:", true)){
						pText = pText.replaceFirst("Keywords:", "");
						Text text = Context.getWmlObjectFactory().createText();
						text.setValue("Keywords:");
						R newR = Context.getWmlObjectFactory().createR();
						newR.getContent().add(text);
						newR.setRPr(Context.getWmlObjectFactory().createRPr());
						newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						newR.getRPr().getRStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.KEYWORDHEADER).getStyleId());
						
						OOXMLConvertTool.removeStringFromPLostRStyle(p, "Keywords:");
						OOXMLConvertTool.removeAllRPROfP(p);
						if(p.getPPr()!=null){
							
						} else {
							p.setPPr(Context.getWmlObjectFactory().createPPr());
						}
						p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.KEYWORDS).getStyleId());
						p.getContent().add(0, newR);
					}
					
					msg.setMessageCode(UserMessage.MESSAGE_KEYWORD_ADAPT_FINISH);
					if(AppEnvironment.getInstance().isActiveTestMode()){
						LogUtils.log(msg.getMessage());
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptLiterature(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(p!=null){
				if(p.getPPr()!=null){
					p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.BIBLIOGRAPHY).getStyleId());
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptLiteratureHeader(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			OOXMLConvertTool.removeAllProperties(p);
			OOXMLConvertTool.removeSectionNumStringOfHeading(p);
			OOXMLConvertTool.convertUppcaseToNormalWord(p);
			p.setPPr(Context.getWmlObjectFactory().createPPr());
			p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
			p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER).getStyleId());
			
			msg.setMessageCode(UserMessage.MESSAGE_BIBLIOGRAPHY_HEADER_ADAPT_FINISH);
			if(AppEnvironment.getInstance().isActiveTestMode()){
				LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptNumberFile(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptPageLayout(MainDocumentPart documentPart, WordprocessingMLPackage wordMLPackage) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(documentPart!=null){
				Body body = wordMLPackage.getMainDocumentPart().getJaxbElement().getBody();
				if(body!=null){
					SectPr bodySectPr = body.getSectPr();
					Style style = springerStyle.get(AcademicFormatStructureDefinition.PAGEFORMAT);
					OOXMLConvertTool.adaptPage(bodySectPr, style, 1);
					
					msg.setMessageCode(UserMessage.MESSAGE_PAGELAYOUT_ADAPT_FINISH);
					if(AppEnvironment.getInstance().isActiveTestMode()){
						LogUtils.log(msg.getMessage());
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptSettings(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptSubtitle(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			OOXMLConvertTool.removePSetting(pPr);
			
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(Context.getWmlObjectFactory().createRPr());
					r.getRPr().setLang(new CTLanguage());
					r.getRPr().getLang().setVal("en-US");
				}
			}
			
			PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
			pStyle.setVal(springerStyle.get(AcademicFormatStructureDefinition.SUBTITLE).getStyleId());
			pPr.setPStyle(pStyle);
			
			msg.setMessageCode(UserMessage.MESSAGE_TITLE_ADAPT_FINISH);
			if(AppEnvironment.getInstance().isActiveTestMode()){
				LogUtils.log(msg.getMessage());
			}
			
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}
	
	@Override
	public UserMessage adaptTitle(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			OOXMLConvertTool.removePSetting(pPr);
			
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(Context.getWmlObjectFactory().createRPr());
					r.getRPr().setLang(new CTLanguage());
					r.getRPr().getLang().setVal("en-US");
				}
			}
			
			PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
			pStyle.setVal(springerStyle.get(AcademicFormatStructureDefinition.TITLE).getStyleId());
			pPr.setPStyle(pStyle);
			
			msg.setMessageCode(UserMessage.MESSAGE_TITLE_ADAPT_FINISH);
			if(AppEnvironment.getInstance().isActiveTestMode()){
				LogUtils.log(msg.getMessage());
			}
			
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptTable(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptTableCaption(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			if(p!=null){
				if(p.getPPr()!=null){
					
				} else {
					p.setPPr(Context.getWmlObjectFactory().createPPr());
				}
				if(p.getPPr().getPStyle()!=null){
					
				} else {
					p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				}
				p.getPPr().getPStyle().setVal(springerStyle.get(AcademicFormatStructureDefinition.CAPTION).getStyleId());
				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
//				OOXMLConvertTool.adaptFirstWord2Special(p, "TABLE", "Table", springerStyle.get(AcademicFormatStructureDefinition.CAPTIONHEADER).getStyleId());
				OOXMLConvertTool.removeFirstStr(p, "TABLE ", ".");
				OOXMLConvertTool.removeFirstStr(p, "Table ", ".");
				
				msg.setMessageCode(UserMessage.MESSAGE_CAPTION_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage appendLayoutPara(MainDocumentPart documentPart, P p, int docContentIndex, int cols) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			Object o = documentPart.getContent().get(docContentIndex+1);
			P newP = null;
			if(o instanceof P){
				String pText = OOXMLConvertTool.getPText((P)o);
				if(pText!=null && pText.replaceAll(" ", "").length()>0){
					newP = Context.getWmlObjectFactory().createP();
					documentPart.getContent().add(docContentIndex+1, newP);
				} else {
					newP = (P)o;
				}
			} else {
				newP = Context.getWmlObjectFactory().createP();
				documentPart.getContent().add(docContentIndex+1, newP);
			}
			newP.setPPr(Context.getWmlObjectFactory().createPPr());
			newP.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
			Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT);
			OOXMLConvertTool.createAppendingLine(newP, style, cols);
			
			msg.setMessageCode(UserMessage.MESSAGE_NEWLINE_LAYOUT_APPENDED);
			if(AppEnvironment.getInstance().isActiveTestMode()){
				LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage removeDefaultHeader(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(documentPart!=null){
				Body body = documentPart.getJaxbElement().getBody();
				if(body!=null){
					SectPr bodySectPr = body.getSectPr();
					if(bodySectPr.getEGHdrFtrReferences()!=null){
						bodySectPr.getEGHdrFtrReferences().clear();
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptGeneratedReferences() {
		// TODO Auto-generated method stub
		return null;
	}

	
}
