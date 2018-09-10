package lib.ooxml.converter;

import java.io.File;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FootnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTFootnotes;
import org.docx4j.wml.CTFtnEdn;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.DocDefaults.RPrDefault;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Text;
import org.docx4j.wml.PPrBase.PStyle;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import lib.ooxml.author.Authors;
import lib.ooxml.identifier.OOXMLIdentifier;
import lib.ooxml.style.OOXMLACMStyle;
import lib.ooxml.tool.OOXMLConvertTool;
import tools.LogUtils;

public class OOXMLConverterACM implements OOXMLConverter {

	private Map<String, Style> acmStyle;
	
	public OOXMLConverterACM() {
		// TODO Auto-generated constructor stub
		this.acmStyle = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap();
	}
	
	@Override
	public UserMessage adaptAbstract(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(p!=null){
				PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
				
				OOXMLConvertTool.removeAllProperties(p);
				
				if(OOXMLConvertTool.isPTextIncludeString(p, "Abstract.", true)	// Springer
						||(OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)8212), true)	// IEEE
								|| OOXMLConvertTool.isPTextIncludeString(p, "Abstract"+((char)45), true))){
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue("Abstract");
					newR.getContent().add(text);
					P newP = Context.getWmlObjectFactory().createP();
					newP.getContent().add(newR);
					newP.setPPr(Context.getWmlObjectFactory().createPPr());
					newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					newP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.ABSTRACTHEADER).getStyleId());
					
					documentPart.getContent().add(docContentIndex, newP);
					
					String s = p.toString();
					s = s.substring(s.indexOf("Abstract")+"Abstract".length()+1, s.length());
					p.getContent().clear();
					
					newR = Context.getWmlObjectFactory().createR();
					text = Context.getWmlObjectFactory().createText();
					text.setValue(s);
					newR.getContent().add(text);
					p.getContent().add(newR);
					
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					pPr.getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.ABSTRACT).getStyleId());
					p.setPPr(pPr);
				} else if(OOXMLConvertTool.isPTextStartWithString(p, "Abstract", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Abstract")) {
					PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
					pStyle.setVal(acmStyle.get(AcademicFormatStructureDefinition.ABSTRACTHEADER).getStyleId());
					pPr.setPStyle(pStyle);
					
					P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
					abstractTextP.setPPr(Context.getWmlObjectFactory().createPPr());
					OOXMLConvertTool.removeAllRPROfP(abstractTextP);
					abstractTextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					abstractTextP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.ABSTRACT).getStyleId());
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
				p.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.FIGURE).getStyleId());
				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
//				OOXMLConvertTool.adaptFirstWord2Special(p, "Fig.", "Figure", null);
				OOXMLConvertTool.removeFirstStr(p, "Fig. ", ".");
				OOXMLConvertTool.removeFirstStr(p, "Figure ", ".");
				
				msg.setMessageCode(UserMessage.MESSAGE_CAPTION_ADAPT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
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
			LogUtils.log(e.getMessage());
		}
		Iterator<Map.Entry<String, Style>> entries = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap().entrySet().iterator();
		
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
				LogUtils.log(e.getMessage());
			}
		}
		
		try {
			List<Object> jaxbNodes = stylePart.getJAXBNodesViaXPath("w:docDefaults", false);
			if(jaxbNodes!=null){
				for (int i = 0; i < jaxbNodes.size(); i++) {
					if(jaxbNodes.get(i) instanceof DocDefaults){
						RPrDefault rPrDefault = (((DocDefaults) jaxbNodes.get(i)).getRPrDefault());
						Style defaultStyle = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap().get(AcademicFormatStructureDefinition.RPRDEFAULT);
						if(defaultStyle!=null && defaultStyle.getRPr()!=null){
							rPrDefault.setRPr(defaultStyle.getRPr());
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
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
										p.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.FOOTNOTE).getStyleId());
										
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
				msg.setMessageCode(UserMessage.MESSAGE_FOOTNOTE_ADAPT_FINISH);
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
				p.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING1).getStyleId());
				
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
				ppr.getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING2).getStyleId());
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
					PPr ppr = OOXMLConvertTool.getPPr(p, true, true);
					if(ppr.getPStyle()!=null){
						
					} else {
						ppr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					}
					ppr.getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING3).getStyleId());
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
				} else {
					
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
					P newP = OOXMLConvertTool.getSameRJudgeWithRRp(p, documentPart.getStyleDefinitionsPart(), true);
					if(newP!=null && newP.getPPr()!=null){
						newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						newP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING3).getStyleId());
						documentPart.getContent().add(docContentIndex, newP);
						p.setPPr(null);
					}
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
					PPr ppr = OOXMLConvertTool.getPPr(p, true, true);
					if(ppr.getPStyle()!=null){
						
					} else {
						ppr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					}
					ppr.getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING4).getStyleId());
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
				} else {
					
					OOXMLConvertTool.removeSectionNumStringOfHeading(p);
					P newP = OOXMLConvertTool.getSameRJudgeWithRRp(p, documentPart.getStyleDefinitionsPart(), true);
					if(newP!=null && newP.getPPr()!=null){
						newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
						newP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING4).getStyleId());
						documentPart.getContent().add(docContentIndex, newP);
						p.setPPr(null);
					}
					
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
				PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
				
				OOXMLConvertTool.removeAllProperties(p);
				
				if(OOXMLConvertTool.isPTextIncludeString(p, "Keywords:", true)	// Springer
						||(OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)8212), true)	// IEEE
						||OOXMLConvertTool.isPTextIncludeString(p, "Keywords-", true)	// IEEE
						|| OOXMLConvertTool.isPTextIncludeString(p, "Keywords"+((char)45), true))){
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue("Keywords");
					newR.getContent().add(text);
					P newP = Context.getWmlObjectFactory().createP();
					newP.getContent().add(newR);
					newP.setPPr(Context.getWmlObjectFactory().createPPr());
					newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					newP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.KEYWORDHEADER).getStyleId());
					
					documentPart.getContent().add(docContentIndex, newP);
					
					String s = p.toString();
					s = s.substring(s.indexOf("Keywords")+"Keywords".length()+1, s.length());
					p.getContent().clear();
					
					newR = Context.getWmlObjectFactory().createR();
					text = Context.getWmlObjectFactory().createText();
					text.setValue(s);
					newR.getContent().add(text);
					p.getContent().add(newR);
					
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					pPr.getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.KEYWORDS).getStyleId());
					p.setPPr(pPr);
				} else if(OOXMLConvertTool.isPTextStartWithString(p, "Keywords", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Keywords")) {
					PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
					pStyle.setVal(acmStyle.get(AcademicFormatStructureDefinition.KEYWORDHEADER).getStyleId());
					pPr.setPStyle(pStyle);
					
					P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
					abstractTextP.setPPr(Context.getWmlObjectFactory().createPPr());
					OOXMLConvertTool.removeAllRPROfP(abstractTextP);
					abstractTextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					abstractTextP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.KEYWORDS).getStyleId());
				}
				msg.setMessageCode(UserMessage.MESSAGE_KEYWORD_ADAPT_FINISH);
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
	public UserMessage adaptLiterature(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		String regex = "(^\\[\\d+\\] \\S.*)|(^\\d+\\. .*)";
		String regs = "(\\[\\d+\\] )|(\\d+\\. )"; 
		String pText = OOXMLConvertTool.getPText(p);
		
		if(pText.matches(regex)){
			OOXMLConvertTool.removeNumStringFromString(p, regs);
		} else {
			
		}
		if(p.getPPr()!=null){
			p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		}
		p.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.BIBLIOGRAPHY).getStyleId());
		
		return msg;
	}

	@Override
	public UserMessage adaptLiteratureHeader(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			OOXMLConvertTool.removeAllProperties(p);
			p.setPPr(Context.getWmlObjectFactory().createPPr());
			p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
			p.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER).getStyleId());
			
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
//					OOXMLConvertTool.fixPageLayout(bodySectPr, "2", "475", false);
//					OOXMLConvertTool.adaptPageLayout(bodySectPr, "2", "720", "12240", "15840", "1", "1080", "1440", "1080", "1080", "720", "720", "0");
					Style style = acmStyle.get(AcademicFormatStructureDefinition.PAGEFORMAT);
					OOXMLConvertTool.adaptPage(bodySectPr, style, 2);
					
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
			pStyle.setVal(acmStyle.get(AcademicFormatStructureDefinition.SUBTITLE).getStyleId());
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
			pStyle.setVal(acmStyle.get(AcademicFormatStructureDefinition.TITLE).getStyleId());
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
				p.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.CAPTION).getStyleId());
				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "TABLE", "Table", null);
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
				} else {
					newP = (P)o;
				}
			}
			newP.setPPr(Context.getWmlObjectFactory().createPPr());
			newP.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
			Style style = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ACM).getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT);
			OOXMLConvertTool.createAppendingLine(newP, style, cols);
			
			documentPart.getContent().add(docContentIndex+1, newP);
			
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
	public UserMessage adaptAcknowledgment(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			if(p!=null){
				
				OOXMLConvertTool.removeSectionNumStringOfHeading(p);
				
				PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
				
				OOXMLConvertTool.removeAllProperties(p);
				
				if(OOXMLConvertTool.isPTextStartWithString(p, "Acknowledgments.", true)){
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue("Acknowledgments");
					newR.getContent().add(text);
					P newP = Context.getWmlObjectFactory().createP();
					newP.getContent().add(newR);
					newP.setPPr(Context.getWmlObjectFactory().createPPr());
					newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					newP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER).getStyleId());
//					newP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING1).getStyleId());
					
					documentPart.getContent().add(docContentIndex, newP);
					
					String s = OOXMLConvertTool.getPText(p);
					s = s.substring(s.indexOf("Acknowledgments")+"Acknowledgments".length()+2, s.length());
					p.getContent().clear();
					
					newR = Context.getWmlObjectFactory().createR();
					text = Context.getWmlObjectFactory().createText();
					text.setValue(s);
					newR.getContent().add(text);
					p.getContent().add(newR);
					
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					pPr.getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.ACKNOWLEDGMENT).getStyleId());
					p.setPPr(pPr);
				} else if(OOXMLConvertTool.isPTextStartWithString(p, "Acknowledgments", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Acknowledgments")	// ACM and Elsevier
						|| OOXMLConvertTool.isPTextStartWithString(p, "Acknowledgment", true) && OOXMLConvertTool.getPText(p).replaceAll(" ", "").equals("Acknowledgment")	// IEEE
						) {
					PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
					pStyle.setVal(acmStyle.get(AcademicFormatStructureDefinition.ACKNOWLEDGMENTHEADER).getStyleId());
//					pStyle.setVal(acmStyle.get(AcademicFormatStructureDefinition.HEADING1).getStyleId());
					pPr.setPStyle(pStyle);
					
					P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
					abstractTextP.setPPr(Context.getWmlObjectFactory().createPPr());
					OOXMLConvertTool.removeAllRPROfP(abstractTextP);
					abstractTextP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					abstractTextP.getPPr().getPStyle().setVal(acmStyle.get(AcademicFormatStructureDefinition.ACKNOWLEDGMENT).getStyleId());
					p.setPPr(pPr);
				}
				
				msg.setMessageCode(UserMessage.MESSAGE_ACKNOWLEDGMENT_ADAPT_FINISH);
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
	public UserMessage adaptAppendixHeader(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptGeneratedReferences() {
		// TODO Auto-generated method stub
		return null;
	}

}
