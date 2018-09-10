package lib.ooxml.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.util.Formatter.BigDecimalLayoutForm;

import javax.xml.bind.JAXBElement;

import org.docx4j.TraversalUtil;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.ClassFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTRel;
import org.docx4j.wml.CTSettings;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STBorder;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STTblLayoutType;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.PStyle;

import base.UserMessage;
import lib.ooxml.author.Author;
import lib.ooxml.author.Authors;
import lib.ooxml.ref.ACMNumbering;
import lib.ooxml.style.OOXMLElsevierStyle2;
import lib.ooxml.tool.OOXMLConvertTool;
import lib.ooxml.tool.OOXMLPageDimension;
import tools.LogUtils;

public class OOXMLElsevier2ConvertingProvider implements OOXMLConvertingInterface {

	private OOXMLElsevierStyle2 elsevierStyle;
	
	private String sourceFileFormat;
	
	private HeaderReference firstHeaderReference;
	
	private HeaderReference defaultHeaderReference;
	
	private HeaderReference evenHeaderReference;
	
	private int titleIndex = 0;
	
	private ArrayList<P> abstractList;
	
	private ArrayList<P> keywordList;
	
	private File convertedFile;
	private WordprocessingMLPackage wordMLPackage;
	private MainDocumentPart documentPart;
	private StyleDefinitionsPart stylePart;
	
	@Override
	public UserMessage adaptAbstract(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			removePSetting(pPr);
			
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(null);
				}
			}
			
			if((OOXMLConvertTool.isPTextStartWithString(p, "Abstract", true)
					|| OOXMLConvertTool.isPTextStartWithString(p, "Abstract.", true))
					&& p.toString().length()>"abstract".length()+2){
				R newR = new R();
				Text text = new Text();
				text.setValue("Abstract");
				newR.getContent().add(text);
				P newP = new P();
				newP.getContent().add(newR);
				newP.setPPr(new PPr());
				newP.getPPr().setPStyle(new PStyle());
				newP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_ABSTRACTHEADER).getStyleId());
				
				documentPart.getContent().add(docContentIndex, newP);
				
				String s = p.toString();
				s = s.substring(s.indexOf("Abstract")+"Abstract".length()+2, s.length());
				p.getContent().clear();
				
				newR = new R();
				text = new Text();
				text.setValue(s);
				newR.getContent().add(text);
				p.getContent().add(newR);
				
				pPr.setPStyle(new PStyle());
				pPr.getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_ABSTRACT).getStyleId());
				p.setPPr(pPr);
				
			} else {
				PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
				pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_ABSTRACTHEADER).getStyleId());
				pPr.setPStyle(pStyle);
				
				P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				abstractTextP.setPPr(new PPr());
				OOXMLConvertTool.removeAllRPROfP(abstractTextP);
				abstractTextP.getPPr().setPStyle(new PStyle());
				abstractTextP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_ABSTRACT).getStyleId());
			}
			
			LogUtils.log("abstract of doc was adapted.");
			
			msg.setMessageCode(UserMessage.MESSAGE_ABSTRACT_ADAPT_FINISH);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptCaption(MainDocumentPart documentPart, P p) {
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
				p.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_CAPTION).getStyleId());
				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "Figure", "Fig.", null);
				
				msg.setMessageCode(UserMessage.MESSAGE_CAPTION_ADAPT_FINISH);
				LogUtils.log(msg.getMessage());
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
		
		int index = 0;
		
		while(index < elsevierStyle.getStyleList().size()){
			try {
				if(OOXMLConvertTool.isStyleExist(elsevierStyle.getStyleList().get(index), stylePart)){
					Style style = stylePart.getStyleById(elsevierStyle.getStyleList().get(index).getStyleId());
					int styleIndex = style!=null?OOXMLConvertTool.getStyleDefinitionIndex(style, stylePart):-1;
					if(styleIndex>=0){
//						stylePart.getContents().getStyle().remove(styleIndex);
						stylePart.getContents().getStyle().set(styleIndex, style);
					}
				} else {
//					documentPart.addStyledParagraphOfText(elsevierStyle.getStyleList().get(index).getStyleId(), "doc "+elsevierStyle.getStyleList().get(index).getName().getVal());
//					stylePart.getContents().getStyle().add(elsevierStyle.getStyleList().get(index));
					ObjectFactory factory = Context.getWmlObjectFactory ();
					Style newStyle = factory.createStyle();
					newStyle.setCustomStyle(true);
					newStyle.setBasedOn(elsevierStyle.getStyleList().get(index).getBasedOn());
					newStyle.setLink(elsevierStyle.getStyleList().get(index).getLink());
					newStyle.setName(elsevierStyle.getStyleList().get(index).getName());
					newStyle.setNext(elsevierStyle.getStyleList().get(index).getNext());
					newStyle.setPPr(elsevierStyle.getStyleList().get(index).getPPr());
					newStyle.setQFormat(elsevierStyle.getStyleList().get(index).getQFormat());
					newStyle.setRPr(elsevierStyle.getStyleList().get(index).getRPr());
					newStyle.setStyleId(elsevierStyle.getStyleList().get(index).getStyleId());
					newStyle.setTblPr(elsevierStyle.getStyleList().get(index).getTblPr());
					newStyle.setTcPr(elsevierStyle.getStyleList().get(index).getTcPr());
					newStyle.setTrPr(elsevierStyle.getStyleList().get(index).getTrPr());
					newStyle.setType(elsevierStyle.getStyleList().get(index).getType());
//					newStyle.setUiPriority(value);
					stylePart.getJaxbElement().getStyle().add(newStyle);
				}
			} catch (Docx4JException e) {
				// TODO Auto-generated catch block
				LogUtils.log("ERROR while adapting style." + e.getMessage());
			}
			index++;
		}
		
		saveChange(convertedFile, wordMLPackage);
		
		LogUtils.log("font class will be modified.");
		
		return msg;
	}

	@Override
	public UserMessage adaptFootNote(File convertedFile, WordprocessingMLPackage wordMLPackage) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage adaptHeading1(P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			removePSetting(pPr);
			
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(new RPr());
					r.getRPr().setLang(new CTLanguage());
					r.getRPr().getLang().setVal("en-US");
				}
			}
			
			PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
			pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_HEADING1).getStyleId());
			pPr.setPStyle(pStyle);
			
			OOXMLConvertTool.removeHeadingNum(1, p);
			
			LogUtils.log("Heading1 of doc was adapted.");
			
			msg.setMessageCode(UserMessage.MESSAGE_HEADING1_ADAPT_FINISH);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading2(P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			removePSetting(pPr);
			
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(new RPr());
					r.getRPr().setLang(new CTLanguage());
					r.getRPr().getLang().setVal("en-US");
				}
			}
			
			PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
			pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_HEADING2).getStyleId());
			pPr.setPStyle(pStyle);
			
			OOXMLConvertTool.removeHeadingNum(2, p);
			
			LogUtils.log("Heading2 of doc was adapted.");
			
			msg.setMessageCode(UserMessage.MESSAGE_HEADING2_ADAPT_FINISH);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading3(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
UserMessage msg = new UserMessage();
		
		try {
			String sectionStr = OOXMLConvertTool.getSectionHeader3String(p, documentPart.getStyleDefinitionsPart());
			if(sectionStr!=null && sectionStr.equals(p.toString())){
				PPr pPr = p.getPPr();
				if(pPr!=null){
					
				} else {
					pPr = Context.getWmlObjectFactory().createPPr();
					p.setPPr(pPr);
				}
				ParaRPr rPr = pPr.getRPr();
				if(rPr!=null){
					
				} else {
					rPr = Context.getWmlObjectFactory().createParaRPr();
					pPr.setRPr(rPr);
				}
				if(rPr.getRStyle()!=null){
					
				} else {
					rPr.setRStyle(Context.getWmlObjectFactory().createRStyle());
				}
				rPr.getRStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_HEADING3).getStyleId());
				P nextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				p.getContent().addAll(nextP.getContent());
				OOXMLConvertTool.removePAt(documentPart, docContentIndex+1);
			} else {
//				String regex = "^\\d\\)\\s";
//				if(adaptHeading3(p, documentPart, docContentIndex, sectionStr, regex)){
//					
//				} else {
//					regex = ".*";
//					adaptHeading3(p, documentPart, docContentIndex, sectionStr, regex);
//				}
			}
			OOXMLConvertTool.removeHeadingNum(3, p);
			
			if(p!=null && p.getPPr()!=null && p.getPPr().getNumPr()!=null){
				p.getPPr().setNumPr(null);
			}
			
			msg.setMessageCode(UserMessage.MESSAGE_HEADING3_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptHeading4(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
UserMessage msg = new UserMessage();
		
		try {
			String sectionStr = OOXMLConvertTool.getSectionHeader3String(p, documentPart.getStyleDefinitionsPart());
			if(sectionStr!=null && sectionStr.equals(p.toString())){
				PPr pPr = p.getPPr();
				if(pPr!=null){
					
				} else {
					pPr = Context.getWmlObjectFactory().createPPr();
					p.setPPr(pPr);
				}
				ParaRPr rPr = pPr.getRPr();
				if(rPr!=null){
					
				} else {
					rPr = Context.getWmlObjectFactory().createParaRPr();
					pPr.setRPr(rPr);
				}
				if(rPr.getRStyle()!=null){
					
				} else {
					rPr.setRStyle(Context.getWmlObjectFactory().createRStyle());
				}
				rPr.getRStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_HEADING4).getStyleId());
				P nextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				p.getContent().addAll(nextP.getContent());
				OOXMLConvertTool.removePAt(documentPart, docContentIndex+1);
			} else {
//				String regex = "^\\d\\)\\s";
//				if(adaptHeading3(p, documentPart, docContentIndex, sectionStr, regex)){
//					
//				} else {
//					regex = ".*";
//					adaptHeading3(p, documentPart, docContentIndex, sectionStr, regex);
//				}
			}
			OOXMLConvertTool.removeHeadingNum(3, p);
			
			if(p!=null && p.getPPr()!=null && p.getPPr().getNumPr()!=null){
				p.getPPr().setNumPr(null);
			}
			
			msg.setMessageCode(UserMessage.MESSAGE_HEADING4_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log("Error in Adapting Heading4. "+msg.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptImages(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		OOXMLConvertTool.fixImgAlignment(documentPart);
		OOXMLConvertTool.fixImgWidth(documentPart);
		msg.setMessageCode(UserMessage.MESSAGE_IMAGE_ADAPT_FINISH);
		LogUtils.log(msg.getMessage());
		return msg;
	}

	@Override
	public UserMessage adaptKeyWord(MainDocumentPart documentPart, P p, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			removePSetting(pPr);
			
			List<Object> pContents = p.getContent();
			for(Object o : pContents){
				if(o instanceof R){
					R r = (R)o;
					r.setRPr(null);
				}
			}
			
			if((OOXMLConvertTool.isPTextStartWithString(p, "Keywords", true)
					|| OOXMLConvertTool.isPTextStartWithString(p, "Keywords.", true))
					&& p.toString().length()>"Keywords".length()+2){
				String s = p.toString();
				s = s.substring("Keywords".length()+2, s.length());
				p.getContent().clear();
				
				R newR = Context.getWmlObjectFactory().createR();
				Text text = Context.getWmlObjectFactory().createText();
				text.setValue("Keywords"+((char)8212));
				OOXMLConvertTool.insertRIntoP(p, newR, 0);
				
				newR = Context.getWmlObjectFactory().createR();
				text = Context.getWmlObjectFactory().createText();
				text.setValue(s);
				OOXMLConvertTool.insertRIntoP(p, newR, 0);
				
				p.setPPr(Context.getWmlObjectFactory().createPPr());
				p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				p.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_KEYWORDHEAD).getStyleId());
				
			} else {
				P keywordTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				keywordTextP.setPPr(new PPr());
				OOXMLConvertTool.removeAllRPROfP(keywordTextP);
				keywordTextP.getPPr().setPStyle(new PStyle());
				keywordTextP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_KEYWORDHEAD).getStyleId());
				
				R newR = Context.getWmlObjectFactory().createR();
				newR.setRPr(Context.getWmlObjectFactory().createRPr());
				newR.getRPr().setI(new BooleanDefaultTrue());
				newR.getRPr().setB(new BooleanDefaultTrue());
				Text text = Context.getWmlObjectFactory().createText();
				text.setValue("Keywords"+((char)8212));
				newR.getContent().add(text);
				OOXMLConvertTool.insertRIntoP(keywordTextP, newR, 0);
				
				OOXMLConvertTool.removePAt(documentPart, docContentIndex);
			}
			
			msg.setMessageCode(UserMessage.MESSAGE_KEYWORD_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptLiterature(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
UserMessage msg = new UserMessage();
		
		try {
			if(p!=null){
				PPr pPr = p.getPPr();
				if(pPr!=null){
					
				} else {
					pPr = Context.getWmlObjectFactory().createPPr();
					p.setPPr(pPr);
				}
				if(pPr.getPStyle()!=null){
					
				} else {
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				}
				pPr.getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_REFERENCE).getStyleId());
				pPr.setInd(new Ind());
				pPr.getInd().setLeft(new BigInteger("360"));
				pPr.getInd().setHanging(new BigInteger("360"));
				
				msg.setMessageCode(UserMessage.MESSAGE_REFERENCE_ADPAT_FINISH);
				LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptLiteratureHeader(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			if(p!=null){
				PPr pPr = p.getPPr();
				if(pPr!=null){
					
				} else {
					pPr = Context.getWmlObjectFactory().createPPr();
					p.setPPr(pPr);
				}
				if(pPr.getPStyle()!=null){
					
				} else {
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				}
				pPr.getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_REFERENCEHEAD).getStyleId());
				msg.setMessageCode(UserMessage.MESSAGE_REFERENCEHEADER_ADPAT_FINISH);
				LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage adaptNumberFile(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		NumberingDefinitionsPart numberingPart = documentPart.getNumberingDefinitionsPart();
		if(numberingPart!=null){
			numberingPart.getAbstractListDefinitions();
		}
		ACMNumbering.adaptNumberingFile(wordMLPackage, documentPart);
		return msg;
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
					bodySectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
					bodySectPr.getCols().setNum(BigInteger.valueOf(1));
					bodySectPr.getCols().setSpace(BigInteger.valueOf(720));
					bodySectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
					bodySectPr.getPgSz().setW(BigInteger.valueOf(11907));
					bodySectPr.getPgSz().setH(BigInteger.valueOf(15876));
					bodySectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
					bodySectPr.getPgMar().setTop(BigInteger.valueOf(714));
					bodySectPr.getPgMar().setBottom(BigInteger.valueOf(879));
					bodySectPr.getPgMar().setLeft(BigInteger.valueOf(1038));
					bodySectPr.getPgMar().setRight(BigInteger.valueOf(947));
					bodySectPr.getPgMar().setHeader(BigInteger.valueOf(907));
					bodySectPr.getPgMar().setFooter(BigInteger.valueOf(1253));
					bodySectPr.getPgMar().setGutter(BigInteger.valueOf(0));
					bodySectPr.setType(Context.getWmlObjectFactory().createSectPrType());
					bodySectPr.getType().setVal("continuous");
					bodySectPr.setTitlePg(new BooleanDefaultTrue());
					bodySectPr.setFootnotePr(Context.getWmlObjectFactory().createCTFtnProps());
					bodySectPr.getFootnotePr().setNumFmt(Context.getWmlObjectFactory().createNumFmt());
					bodySectPr.getFootnotePr().getNumFmt().setVal(NumberFormat.CHICAGO);
//					OOXMLConvertTool.fixPageLayout(bodySectPr, "1", "720", false);
//					OOXMLPageDimension pd = new OOXMLPageDimension(bodySectPr);
					
					/*
					 * fix: default header need to remove
					 */
					if(bodySectPr.getEGHdrFtrReferences()!=null){
						bodySectPr.getEGHdrFtrReferences().clear();
						bodySectPr.getEGHdrFtrReferences().add(defaultHeaderReference);
						bodySectPr.getEGHdrFtrReferences().add(firstHeaderReference);
					}
					
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptTitle(MainDocumentPart documentPart, P p) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
			
			removePSetting(pPr);
			
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
			pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_TITLE).getStyleId());
			pPr.setPStyle(pStyle);
			
			LogUtils.log("Title of doc was adapted.");
			
			msg.setMessageCode(UserMessage.MESSAGE_TITLE_ADAPT_FINISH);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		
		return msg;
	}

	@Override
	public UserMessage adaptTable(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
UserMessage msg = new UserMessage();
		
		try {
			ClassFinder classFinder = new ClassFinder(Tbl.class);
			new TraversalUtil(documentPart.getContent(), classFinder);
			int index = 0;
			while(index<classFinder.results.size()){
				Tbl tbl = (Tbl) classFinder.results.get(index);
				if(tbl.getTblPr()!=null){
//					tbl.getTblPr().getTblCaption();
					if(tbl.getTblPr().getTblW()!=null
							&& tbl.getTblPr().getTblW().getType()!=null
							&& !tbl.getTblPr().getTblW().getType().equals("auto")
							&& tbl.getTblPr().getTblW().getW()!=null
							&& tbl.getTblPr().getTblW().getW().intValue()>12240/2
							&& tbl.getTblGrid()!=null
							&& tbl.getTblGrid().getGridCol()!=null
							&& tbl.getTblGrid().getGridCol().size()<6){
//						System.out.println("something in 1");
						tbl.getTblPr().getTblW().setW(new BigInteger(""+(12240/2)));
						tbl.getTblPr().setTblLayout(Context.getWmlObjectFactory().createCTTblLayoutType());
						tbl.getTblPr().getTblLayout().setType(STTblLayoutType.AUTOFIT);
					} else if(tbl.getTblGrid()!=null
							&& tbl.getTblGrid().getGridCol()!=null
							&& tbl.getTblGrid().getGridCol().size()<6){
						int calTblW = 0;
						for(TblGridCol col : tbl.getTblGrid().getGridCol()){
							if(col.getW()!=null){
								calTblW += col.getW().intValue();
							}
						}
						int calDiff = (12240/2)-calTblW;
						int calFactor = 0;
						if(calDiff<0){
							calFactor = calDiff/tbl.getTblGrid().getGridCol().size();
							for(TblGridCol col : tbl.getTblGrid().getGridCol()){
								if(col.getW()!=null){
									col.setW(new BigInteger((col.getW().intValue()+calFactor)+""));
								}
							}
//							System.out.println("something in 2");
						}
						
					} else if(tbl.getTblGrid()!=null
							&& tbl.getTblGrid().getGridCol()!=null
							&& tbl.getTblGrid().getGridCol().size()>=6){
						tbl.getTblPr().setTblLayout(Context.getWmlObjectFactory().createCTTblLayoutType());
						tbl.getTblPr().getTblLayout().setType(STTblLayoutType.AUTOFIT);
//						System.out.println("something in 3");
					} else {
//						System.out.println("something in 4");
					}
				}
				index++;
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return msg;
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
				p.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_CAPTION).getStyleId());
//				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setRPr(null);
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "TABLE", "Table", null);
				
				msg.setMessageCode(UserMessage.MESSAGE_TABLECAPTION_ADAPT_FINISH);
				LogUtils.log(msg.getMessage());
			}
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
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
			if(o instanceof P && ((P)o).toString().length()>0){
				newP = Context.getWmlObjectFactory().createP();
			} else {
				newP = (P)o;
			}
			
			newP.setPPr(Context.getWmlObjectFactory().createPPr());
			newP.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
			newP.getPPr().getSectPr().setCols(Context.getWmlObjectFactory().createCTColumns());
			SectPr sectPr = newP.getPPr().getSectPr();
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(cols));
			sectPr.getCols().setSpace(BigInteger.valueOf(720));
			sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
			sectPr.getPgSz().setW(BigInteger.valueOf(11907));
			sectPr.getPgSz().setH(BigInteger.valueOf(15876));
			sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
			sectPr.getPgMar().setTop(BigInteger.valueOf(714));
			sectPr.getPgMar().setBottom(BigInteger.valueOf(879));
			sectPr.getPgMar().setLeft(BigInteger.valueOf(1038));
			sectPr.getPgMar().setRight(BigInteger.valueOf(947));
			sectPr.getPgMar().setHeader(BigInteger.valueOf(907));
			sectPr.getPgMar().setFooter(BigInteger.valueOf(1253));
			sectPr.getPgMar().setGutter(BigInteger.valueOf(0));
			sectPr.setType(Context.getWmlObjectFactory().createSectPrType());
			sectPr.getType().setVal("continuous");
			sectPr.setTitlePg(new BooleanDefaultTrue());
			sectPr.setFootnotePr(Context.getWmlObjectFactory().createCTFtnProps());
			sectPr.getFootnotePr().setNumFmt(Context.getWmlObjectFactory().createNumFmt());
			sectPr.getFootnotePr().getNumFmt().setVal(NumberFormat.CHICAGO);
			sectPr.getEGHdrFtrReferences().add(firstHeaderReference);
			documentPart.getContent().add(docContentIndex+1, newP);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return msg;
	}
	public UserMessage appendLayoutPara2(MainDocumentPart documentPart, int docContentIndex) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		try {
			P newP = Context.getWmlObjectFactory().createP();
			
			newP.setPPr(Context.getWmlObjectFactory().createPPr());
			newP.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
			newP.getPPr().getSectPr().setCols(Context.getWmlObjectFactory().createCTColumns());
			SectPr sectPr = newP.getPPr().getSectPr();
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(2));
			sectPr.getCols().setSpace(BigInteger.valueOf(720));
			sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
			sectPr.getPgSz().setW(BigInteger.valueOf(11907));
			sectPr.getPgSz().setH(BigInteger.valueOf(15876));
			sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
			sectPr.getPgMar().setTop(BigInteger.valueOf(714));
			sectPr.getPgMar().setBottom(BigInteger.valueOf(879));
			sectPr.getPgMar().setLeft(BigInteger.valueOf(1038));
			sectPr.getPgMar().setRight(BigInteger.valueOf(947));
			sectPr.getPgMar().setHeader(BigInteger.valueOf(907));
			sectPr.getPgMar().setFooter(BigInteger.valueOf(1253));
			sectPr.getPgMar().setGutter(BigInteger.valueOf(0));
			sectPr.setType(Context.getWmlObjectFactory().createSectPrType());
			sectPr.getType().setVal("continuous");
			sectPr.setTitlePg(new BooleanDefaultTrue());
			sectPr.setFootnotePr(Context.getWmlObjectFactory().createCTFtnProps());
			sectPr.getFootnotePr().setNumFmt(Context.getWmlObjectFactory().createNumFmt());
			sectPr.getFootnotePr().getNumFmt().setVal(NumberFormat.CHICAGO);
			sectPr.getEGHdrFtrReferences().add(defaultHeaderReference);
			documentPart.getContent().add(docContentIndex+1, newP);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return msg;
	}
	@Override
	public UserMessage convert2Format(File convertedFile, WordprocessingMLPackage wordMLPackage,
			MainDocumentPart documentPart, StyleDefinitionsPart stylePart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			OOXMLConvertTool.sourceFormat = null;
			List<Object> docContentList = documentPart.getContent();
			elsevierStyle = new OOXMLElsevierStyle2();
			abstractList = new ArrayList<P>();
			keywordList = new ArrayList<P>();
			
			Authors authors = new Authors();
			
			adaptFormatHeader(wordMLPackage);
			adaptPageLayout(documentPart, wordMLPackage);
			
			int docContentIndex = 0;
			int authorCount = 0;
			boolean isAuthorbeginn = false;
			boolean isKeywordBegin = false;
			boolean isAbstractBegin = false;
			int authorInfoLine = 0;
			int authorNodeCount = 0;
			
			while(docContentIndex<docContentList.size()){
				Object o = docContentList.get(docContentIndex);
				if(o instanceof P) {
					P p = (P) o;
//					if(OOXMLConvertTool.isImage(p)){
//						System.out.println("test is image");
//					} else 
					if(OOXMLConvertTool.isTitle(p, stylePart)){
						titleIndex = docContentIndex;
						adaptTitle(documentPart, p);
						appendLayoutPara(documentPart, p, docContentIndex, 1);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isAuthors(p, stylePart)){
						authorNodeCount++;
						authorInfoLine++;
						if(!OOXMLConvertTool.isEmail(p)){
							authorInfoLine = 0;
							isAuthorbeginn=false;
						}
						if(!isAuthorbeginn){
							isAuthorbeginn=true;
							authorCount++;
							authors.getAuthors().add(new Author());
							authors.getAuthors().get(authorCount-1).setName(p.toString());
						} else {
							if(OOXMLConvertTool.isEmail(p)){
								authors.getAuthors().get(authorCount-1).setInfo4(p.toString());
								authorInfoLine = 0;
								isAuthorbeginn=false;
							}
						}
					} else if(OOXMLConvertTool.isAbstract(p, stylePart)) {
						isAbstractBegin = true;
						isKeywordBegin = false;
						isAuthorbeginn=false;
						
						adaptAuthors(documentPart, docContentIndex, authorNodeCount, authors);
						docContentIndex = docContentIndex-authorNodeCount;
						docContentIndex = docContentIndex + authors.getAuthors().size();
						adaptAbstract(documentPart, p, docContentIndex);
						
						abstractList.add(p);
						OOXMLConvertTool.removePAt(documentPart, docContentIndex);
						docContentIndex--;
						
						
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isKeywords(p, stylePart)) {
						isKeywordBegin = true;
						isAbstractBegin = false;
						isAuthorbeginn=false;
						
						adaptKeyWord(documentPart, p, docContentIndex);
						
						keywordList.add(p);
						OOXMLConvertTool.removePAt(documentPart, docContentIndex);
						docContentIndex--;
						
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isSectionHeader1(p, stylePart, documentPart)){
						if(isAbstractBegin||isKeywordBegin){
							if(abstractList!=null && abstractList.size()>0){
								documentPart.getContent().add(docContentIndex, createAbstractTbl());
							}
						}
						isAbstractBegin = false;
						isKeywordBegin = false;
						isAuthorbeginn=false;
						adaptHeading1(p);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isSectionHeader2(p, stylePart, documentPart)) {
						isAbstractBegin = false;
						isKeywordBegin = false;
						isAuthorbeginn=false;
						adaptHeading2(p);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isSectionHeader3(p, stylePart, documentPart)) {
						adaptHeading3(documentPart, p, docContentIndex);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isSectionHeader4(p, stylePart, documentPart)) {
						adaptHeading4(documentPart, p, docContentIndex);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isCaption(p, documentPart, stylePart)){
						adaptCaption(documentPart, p);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isTableCaption(p, stylePart)){
						adaptTableCaption(documentPart, p);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isAcknowledgReferenceHeader(p, stylePart, documentPart)){
						adaptLiteratureHeader(documentPart, p);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isReferenceText(p, stylePart, documentPart)){
						adaptLiterature(documentPart, p);
						saveChange(convertedFile, wordMLPackage);
					} else {
						removePSetting(p.getPPr());
						if(isAuthorbeginn){
							authorInfoLine++;
							authors.getAuthors().get(authorCount-1).setInfo(authorInfoLine, p.toString());
							authorNodeCount++;
						}
						if(isKeywordBegin){
							keywordList.add(p);
							OOXMLConvertTool.removePAt(documentPart, docContentIndex);
							docContentIndex--;
						}
						if(isAbstractBegin){
							abstractList.add(p);
							OOXMLConvertTool.removePAt(documentPart, docContentIndex);
							docContentIndex--;
						}
					}
					if(docContentIndex+1>=docContentList.size()){
						appendLayoutPara2(documentPart, docContentIndex);
						System.out.println("done");
						docContentIndex++;
					}
				} else if(o instanceof Tbl){
					Tbl tbl = (Tbl)o;
				} else if(o instanceof JAXBElement<?>){
					System.out.println(((JAXBElement) o).getValue().getClass().getName());
					System.out.println(docContentIndex+":"+docContentList.size());
				} else if(docContentIndex+1>=docContentList.size()){
					appendLayoutPara2(documentPart, docContentIndex);
					System.out.println("done");
				} else {
					System.out.println(o.getClass().getName());
				}
				docContentIndex++;
			}
			
			adaptFootNote(convertedFile, wordMLPackage);
			adaptImages(documentPart);
			adaptFontStyleClass(convertedFile, wordMLPackage, documentPart, stylePart);
			
			adaptNumberFile(documentPart);
			adaptTable(documentPart);
			
			System.out.println(abstractList.size()+"##############");
			
			saveChange(convertedFile, wordMLPackage);
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		msg.setMessageCode(UserMessage.MESSAGE_GENERAL_OK);
		
		return msg;
	}

	@Override
	public UserMessage removePSetting(PPr pPr) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
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
		
		return msg;
	}

	@Override
	public UserMessage removeUselessPageBreaks(WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public UserMessage saveChange(File convertedFile, WordprocessingMLPackage wordMLPackage) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			wordMLPackage.save(convertedFile);
			msg.setMessageDetails("The Changes have been saved.");
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage adaptAuthors(MainDocumentPart documentPart, int docContentIndex, int authorNodeCount, Authors authors) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		if(authorNodeCount>0){
			int index = docContentIndex-authorNodeCount-2;
			while(index<=titleIndex){
				index++;
			}
			for(int i=0; i<=authorNodeCount; i++){
				documentPart.getContent().remove(index);
			}
			
			P newP = Context.getWmlObjectFactory().createP();
			newP.setPPr(Context.getWmlObjectFactory().createPPr());
			newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
			newP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_AUTHORS).getStyleId());
			R r = Context.getWmlObjectFactory().createR();
			Text text = Context.getWmlObjectFactory().createText();
			text.setValue(authors.getNameString());
			
			r.getContent().add(text);
			newP.getContent().add(r);
			documentPart.getContent().add(index, newP);
			int i=0;
			for(Author author: authors.getAuthors()){
				newP = Context.getWmlObjectFactory().createP();
				newP.setPPr(Context.getWmlObjectFactory().createPPr());
				newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				newP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle2.STYLENAME_AFFILIATIONS).getStyleId());
				r = Context.getWmlObjectFactory().createR();
				text = Context.getWmlObjectFactory().createText();
				text.setValue(author.getInfo());
				
				r.getContent().add(text);
				newP.getContent().add(r);
				documentPart.getContent().add(index+(++i), newP);
			}
			
			msg.setMessageCode(UserMessage.MESSAGE_AUTHORS_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
		}
		return msg;
	}

	private UserMessage adaptFormatHeader(WordprocessingMLPackage wordMLPackage){
		UserMessage msg = new UserMessage();
		
		try {
			removeExistsHeaderReference(wordMLPackage);
			createDefaultHeaderPart(wordMLPackage);
			createHeaderPart(wordMLPackage);
//			OOXMLElsevierHeaderFooterProvider.createHeader(wordMLPackage);
			
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}

		return msg;
	}

	private void removeExistsHeaderReference(WordprocessingMLPackage wordMLPackage){
		List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
		SectPr sectionProperties = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectionProperties == null) {
			ObjectFactory factory = Context.getWmlObjectFactory();
			sectionProperties = factory.createSectPr();
			wordMLPackage.getMainDocumentPart().addObject(sectionProperties);
			sections.get(sections.size() - 1).setSectPr(sectionProperties);
		}

		/*
		 * Remove Header if it already exists.
		 */
		List<CTRel> relations = sectionProperties.getEGHdrFtrReferences();
		Iterator<CTRel> relationsItr = relations.iterator();
		while (relationsItr.hasNext()) {
			CTRel relation = relationsItr.next();
			if (relation instanceof HeaderReference) {
				relationsItr.remove();
			}
		}
		
	}
	
	private void createDefaultHeaderPart(WordprocessingMLPackage wordMLPackage) {
		try {
			ObjectFactory factory = Context.getWmlObjectFactory();
			PartName partName = new PartName("/word/header/header2.xml");
			HeaderPart headerPart = new HeaderPart(partName);
			headerPart.setPackage(wordMLPackage);

			P p = Context.getWmlObjectFactory().createP();
			R r = Context.getWmlObjectFactory().createR();
			Text text = Context.getWmlObjectFactory().createText();
			p.getContent().add(r);
			p.setPPr(Context.getWmlObjectFactory().createPPr());
			p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
			p.getPPr().getJc().setVal(JcEnumeration.CENTER);
			r.getContent().add(text);
			r.setRPr(Context.getWmlObjectFactory().createRPr());
			r.getRPr().setI(new BooleanDefaultTrue());
			r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
			r.getRPr().getSz().setVal(new BigInteger("16"));
			r.getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
			r.getRPr().getRFonts().setAscii("Times New Roman");
			text.setValue("Author name / Procedia Computer Science 000 (2017) 000–111");
			
			Hdr header = factory.createHdr();
			header.getContent().add(p);
			headerPart.setJaxbElement(header);

			Relationship relationship = wordMLPackage.getMainDocumentPart().addTargetPart(headerPart);
//			relationship.setTarget("header/elsevier_header.xml");
			
			List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
			SectPr sectionProperties = sections.get(sections.size() - 1).getSectPr();
			// There is always a section wrapper, but it might not contain a sectPr
			if (sectionProperties == null) {
				sectionProperties = factory.createSectPr();
				wordMLPackage.getMainDocumentPart().addObject(sectionProperties);
				sections.get(sections.size() - 1).setSectPr(sectionProperties);
			}
			sectionProperties.setTitlePg(new BooleanDefaultTrue());
			
			HeaderReference headerReference = factory.createHeaderReference();
			headerReference.setId(relationship.getId());
			headerReference.setType(HdrFtrRef.DEFAULT);
			sectionProperties.getEGHdrFtrReferences().add(headerReference);
			defaultHeaderReference = headerReference;
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
	}

	
	private void createHeaderPart(WordprocessingMLPackage wordMLPackage) {
		try {
			ObjectFactory factory = Context.getWmlObjectFactory();
			PartName partName = new PartName("/word/header/header1.xml");
			HeaderPart headerPart = new HeaderPart(partName);
			headerPart.setPackage(wordMLPackage);

//			Hdr header = getHdr(wordMLPackage, headerPart);
			Hdr header = factory.createHdr();
			
			header.getContent().add(createDocHeaderTbl(wordMLPackage, headerPart));
			headerPart.setJaxbElement(header);

			Relationship relationship = wordMLPackage.getMainDocumentPart().addTargetPart(headerPart);
//			relationship.setTarget("header/elsevier_header.xml");
			List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();
			SectPr sectionProperties = sections.get(sections.size() - 1).getSectPr();
			// There is always a section wrapper, but it might not contain a sectPr
			if (sectionProperties == null) {
				sectionProperties = factory.createSectPr();
				wordMLPackage.getMainDocumentPart().addObject(sectionProperties);
				sections.get(sections.size() - 1).setSectPr(sectionProperties);
			}
			sectionProperties.setTitlePg(new BooleanDefaultTrue());
			
			HeaderReference firstPageHeaderRef = factory.createHeaderReference();
			firstPageHeaderRef.setId(relationship.getId());
			firstPageHeaderRef.setType(HdrFtrRef.FIRST);
			sectionProperties.getEGHdrFtrReferences().add(firstPageHeaderRef);
			firstHeaderReference = firstPageHeaderRef;
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
		
	}
	
	private Tbl createDocHeaderTbl(WordprocessingMLPackage wordMLPackage, HeaderPart headerPart){
		Tbl tbl = null;
		
		try {
			int writableWidthTwips = wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
			tbl = TblFactory.createTable(1, 3, writableWidthTwips);// in 1 row, 3 columns, auto width
			
			tbl.setTblGrid(Context.getWmlObjectFactory().createTblGrid());
			
			int[] colWidth = {1200, writableWidthTwips-1200-1920, 1920};
			
			TblGridCol tblCol = Context.getWmlObjectFactory().createTblGridCol();
			tblCol.setW(new BigInteger(""+colWidth[0]));
			tbl.getTblGrid().getGridCol().add(tblCol);
			tblCol = Context.getWmlObjectFactory().createTblGridCol();
			tblCol.setW(new BigInteger(""+colWidth[1]));
			tbl.getTblGrid().getGridCol().add(tblCol);
			tblCol = Context.getWmlObjectFactory().createTblGridCol();
			tblCol.setW(new BigInteger(""+colWidth[2]));
			tbl.getTblGrid().getGridCol().add(tblCol);
			
			int i=0;
			
			for(Object o : tbl.getContent()) {
				if(o instanceof Tr){
					Tr tr = (Tr)o;
					for (Object trContentObject : tr.getContent()) {
						if(trContentObject instanceof Tc){
							Tc tc = (Tc) trContentObject;
							tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
							tc.getTcPr().setVAlign(Context.getWmlObjectFactory().createCTVerticalJc());
							tc.getTcPr().getVAlign().setVal(STVerticalJc.CENTER);
							
							P p;
							i++;
							URL imageURL;
							File imageFile;
							byte[] bytes;
							switch (i) {
							case 1:
								tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
								tc.getTcPr().setTcW(Context.getWmlObjectFactory().createTblWidth());
								tc.getTcPr().getTcW().setW(new BigInteger(""+colWidth[0]));
								tc.getTcPr().getTcW().setType("dxa");
								
								imageURL = getClass().getResource("/sources/elsevier/elsevier_header_image1.png");
								imageFile = new File(imageURL.toURI());
								
//								imageFile = new File("src/sources/elsevier/elsevier_header_image1.png");
								bytes = convertImageToByteArray(imageFile);
								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Elsevier", "Elsevier logo", 604800, 669600);
								tc.getContent().add(p);
								break;
							case 2:
								tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
								tc.getTcPr().setTcW(Context.getWmlObjectFactory().createTblWidth());
								tc.getTcPr().getTcW().setW(new BigInteger(""+colWidth[1]));
								tc.getTcPr().getTcW().setType("dxa");
								// 1
								p = Context.getWmlObjectFactory().createP();
								p.setPPr(Context.getWmlObjectFactory().createPPr());
								p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
								p.getPPr().getJc().setVal(JcEnumeration.CENTER);
								p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
								p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//								p.getPPr().getSpacing().setAfter(new BigInteger("240"));
//								p.getPPr().getSpacing().setBefore(new BigInteger("100"));
								p.getPPr().getSpacing().setLine(new BigInteger("200"));
								p.getPPr().getSpacing().setLineRule(STLineSpacingRule.AT_LEAST);
								
								R r = Context.getWmlObjectFactory().createR();
								r.setRPr(Context.getWmlObjectFactory().createRPr());
								r.getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
								r.getRPr().getRFonts().setAscii("Arial");
								r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
								r.getRPr().getSz().setVal(new BigInteger("18"));
								Text text = Context.getWmlObjectFactory().createText();
								text.setValue("Available online at ");
								r.getContent().add(text);
								p.getContent().add(r);
								
//								Hyperlink hyperLink = createHyperlink(wordMLPackage.getMainDocumentPart(), "http://www.sciencedirect.com/");
//								p.getContent().add(hyperLink);
								r = Context.getWmlObjectFactory().createR();
								text = Context.getWmlObjectFactory().createText();
								text.setValue(" http://www.sciencedirect.com/");
								r.getContent().add(text);
								p.getContent().add(r);
								
								tc.getContent().add(p);
								// 2
								p = Context.getWmlObjectFactory().createP();
								p.setPPr(Context.getWmlObjectFactory().createPPr());
								p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
								p.getPPr().getJc().setVal(JcEnumeration.CENTER);
								p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
								p.getPPr().getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
								p.getPPr().getRPr().getRFonts().setAscii("Arial");
								p.getPPr().getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
								p.getPPr().getRPr().getSz().setVal(new BigInteger("33"));
								p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//								p.getPPr().getSpacing().setAfter(new BigInteger("240"));
//								p.getPPr().getSpacing().setBefore(new BigInteger("100"));
								p.getPPr().getSpacing().setLine(new BigInteger("300"));
								p.getPPr().getSpacing().setLineRule(STLineSpacingRule.AT_LEAST);
								
								r = Context.getWmlObjectFactory().createR();
								r.setRPr(Context.getWmlObjectFactory().createRPr());
								r.getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
								r.getRPr().getRFonts().setAscii("Arial");
								r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
								r.getRPr().getSz().setVal(new BigInteger("33"));
								text = Context.getWmlObjectFactory().createText();
								text.setValue("ScienceDirect");
								r.getContent().add(text);
								p.getContent().add(r);
								
								tc.getContent().add(p);
								// 3
								p = Context.getWmlObjectFactory().createP();
								p.setPPr(Context.getWmlObjectFactory().createPPr());
								p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
								p.getPPr().getJc().setVal(JcEnumeration.CENTER);
								p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//								p.getPPr().getSpacing().setAfter(new BigInteger("240"));
//								p.getPPr().getSpacing().setBefore(new BigInteger("100"));
								p.getPPr().getSpacing().setLine(new BigInteger("200"));
								p.getPPr().getSpacing().setLineRule(STLineSpacingRule.AT_LEAST);
								
								r = Context.getWmlObjectFactory().createR();
								r.setRPr(Context.getWmlObjectFactory().createRPr());
								r.getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
								r.getRPr().getRFonts().setAscii("Times New Roman");
								r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
								r.getRPr().getSz().setVal(new BigInteger("16"));
								text = Context.getWmlObjectFactory().createText();
								text.setValue("Procedia Computer Science 000 (2017) 000–111");
								r.getContent().add(text);
								p.getContent().add(r);
								
								tc.getContent().add(p);
								break;
							case 3:
								tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
								tc.getTcPr().setTcW(Context.getWmlObjectFactory().createTblWidth());
								tc.getTcPr().getTcW().setW(new BigInteger(""+colWidth[2]));
								tc.getTcPr().getTcW().setType("dxa");
								
								imageURL = getClass().getResource("/sources/elsevier/elsevier_header_cs_image3.png");
								imageFile = new File(imageURL.toURI());
								bytes = convertImageToByteArray(imageFile);
								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Procedia", "Procedia logo", 1137600, 741600);
								
								tc.getContent().add(p);
							default:
								break;
							}
							
						} else {
							System.out.println(trContentObject.getClass().getName());
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
		
		return tbl;
	}
	
	private static byte[] convertImageToByteArray(File file) {
		
		byte[] bytes = null;
		try {
			InputStream is = new FileInputStream(file);
			long length = file.length();
			// You cannot create an array using a long, it needs to be an int.
			if (length > Integer.MAX_VALUE) {
				System.out.println("File too large!!");
			}
			bytes = new byte[(int) length];
			int offset = 0;
			int numRead = 0;
			while (offset < bytes.length && (numRead = is.read(bytes, offset, bytes.length - offset)) >= 0) {
				offset += numRead;
			}
			// Ensure all the bytes have been read
			if (offset < bytes.length) {
				System.out.println("Could not completely read file " + file.getName());
			}
			is.close();
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
		
		return bytes;
		
	}
	private P addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes, HeaderPart part, String fileNameHint, String altText, long cx, long cy) {
		P paragraph = null;
		try {
			BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, part, bytes);

			int docPrId = 1;
			int cNvPrId = 2;
			Inline inline = imagePart.createImageInline(fileNameHint, altText, docPrId, cNvPrId, false);
			if(cx>0 && cy>0){
				inline.getExtent().setCx(cx);
				inline.getExtent().setCy(cy);
			}
			
			paragraph = addInlineImageToParagraph(inline);

//			wordMLPackage.getMainDocumentPart().addObject(paragraph);

		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
		return paragraph;
	}

	private P addInlineImageToParagraph(Inline inline) {
		// Now add the in-line image to a paragraph
		ObjectFactory factory = Context.getWmlObjectFactory();
		P paragraph = factory.createP();
		R run = factory.createR();
		paragraph.getContent().add(run);
		Drawing drawing = factory.createDrawing();
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		return paragraph;
	}
	
	private Tbl createAbstractTbl(){
		ArrayList<P> oldAbstractPs = getAbstract();
		Tbl tbl = Context.getWmlObjectFactory().createTbl();
		
		tbl.setTblPr(Context.getWmlObjectFactory().createTblPr());
		tbl.getTblPr().setTblW(Context.getWmlObjectFactory().createTblWidth());
		tbl.getTblPr().getTblW().setW(BigInteger.valueOf(9923));
		tbl.getTblPr().getTblW().setType("dxa");
		
		tbl.getTblPr().setTblInd(Context.getWmlObjectFactory().createTblWidth());
		tbl.getTblPr().getTblInd().setW(BigInteger.valueOf(108));
		tbl.getTblPr().getTblInd().setType("dxa");
		
		tbl.getTblPr().setTblLayout(Context.getWmlObjectFactory().createCTTblLayoutType());
		tbl.getTblPr().getTblLayout().setType(STTblLayoutType.FIXED);
		
		tbl.getTblPr().setTblBorders(Context.getWmlObjectFactory().createTblBorders());
		CTBorder border = Context.getWmlObjectFactory().createCTBorder();
		border.setVal(STBorder.NONE);
		tbl.getTblPr().getTblBorders().setBottom(border);
		tbl.getTblPr().getTblBorders().setTop(border);
		tbl.getTblPr().getTblBorders().setLeft(border);
		tbl.getTblPr().getTblBorders().setRight(border);
		tbl.getTblPr().getTblBorders().setInsideH(border);
		tbl.getTblPr().getTblBorders().setInsideV(border);
		
		tbl.setTblGrid(Context.getWmlObjectFactory().createTblGrid());
		TblGridCol col = Context.getWmlObjectFactory().createTblGridCol();
		col.setW(BigInteger.valueOf(2662));
		tbl.getTblGrid().getGridCol().add(col);
		col = Context.getWmlObjectFactory().createTblGridCol();
		col.setW(BigInteger.valueOf(634));
		tbl.getTblGrid().getGridCol().add(col);
		col = Context.getWmlObjectFactory().createTblGridCol();
		col.setW(BigInteger.valueOf(6627));
		tbl.getTblGrid().getGridCol().add(col);
		
		int index = 0;
		for(Object o : tbl.getContent()) {
			if(o instanceof Tr){
				Tr tr = (Tr)o;
				for (Object trContentObject : tr.getContent()) {
					if(trContentObject instanceof Tc){
						Tc tc = (Tc) trContentObject;
						tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
						tc.getTcPr().setVAlign(Context.getWmlObjectFactory().createCTVerticalJc());
						tc.getTcPr().getVAlign().setVal(STVerticalJc.CENTER);
						
						
						switch (index) {
						case 0:
							tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
							tc.getTcPr().setTcW(Context.getWmlObjectFactory().createTblWidth());
							tc.getTcPr().getTcW().setW(BigInteger.valueOf(2662));
							tc.getTcPr().getTcW().setType("dxa");
							
							P p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							R r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							Text text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(18));
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setBefore(BigInteger.valueOf(440));
							p.getPPr().getSpacing().setAfter(BigInteger.valueOf(200));
							text.setValue("A R T I C L E  I N F O");
							
							p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setI(new BooleanDefaultTrue());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(15));
							
							p.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
							CTBorder border1 = Context.getWmlObjectFactory().createCTBorder();
							border1.setColor("000000");
							border1.setVal(STBorder.SINGLE);
							border1.setSz(BigInteger.valueOf(4));
							border1.setSpace(BigInteger.valueOf(1));
							p.getPPr().getPBdr().setTop(border1);
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setLine(BigInteger.valueOf(230));
							p.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
							text.setValue("Article history:");
							
							p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(15));
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setLine(BigInteger.valueOf(230));
							p.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
							text.setValue("Received 00 December 00");
							
							p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(15));
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setLine(BigInteger.valueOf(230));
							p.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
							text.setValue("Received in revised form 00 January 00");
							
							p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(15));
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setLine(BigInteger.valueOf(230));
							p.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
							text.setValue("Accepted 00 February 00");
							
							p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setI(new BooleanDefaultTrue());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(15));
							p.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
							border1 = Context.getWmlObjectFactory().createCTBorder();
							p.getPPr().getPBdr().setTop(border1);
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setLine(BigInteger.valueOf(230));
							p.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
							text.setValue("Keywords:");
							
							p = Context.getWmlObjectFactory().createP();
							tc.getContent().add(p);
							r = Context.getWmlObjectFactory().createR();
							p.getContent().add(r);
							text = Context.getWmlObjectFactory().createText();
							r.getContent().add(text);
							r.setRPr(Context.getWmlObjectFactory().createRPr());
							r.getRPr().setI(new BooleanDefaultTrue());
							r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
							r.getRPr().getSz().setVal(BigInteger.valueOf(15));
							p.getPPr().setPBdr(Context.getWmlObjectFactory().createPPrBasePBdr());
							p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
							p.getPPr().getSpacing().setLine(BigInteger.valueOf(230));
							p.getPPr().getSpacing().setLineRule(STLineSpacingRule.EXACT);
							text.setValue("Each keyword to start on a new line ");
							break;
						case 1:
							tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
							tc.getTcPr().setTcW(Context.getWmlObjectFactory().createTblWidth());
							tc.getTcPr().getTcW().setW(BigInteger.valueOf(634));
							tc.getTcPr().getTcW().setType("dxa");
							break;
						case 2:
							tc.setTcPr(Context.getWmlObjectFactory().createTcPr());
							tc.getTcPr().setTcW(Context.getWmlObjectFactory().createTblWidth());
							tc.getTcPr().getTcW().setW(BigInteger.valueOf(6627));
							tc.getTcPr().getTcW().setType("dxa");
							for(P oldP : oldAbstractPs){
								tc.getContent().add(oldP);
							}
							break;
						default:
							break;
						}
					}
				}
			}
		}
		return tbl;
	}
	public void setAbstract(ArrayList<P> abstractList){
		this.abstractList = abstractList;
	}
	public ArrayList<P> getAbstract(){
		return abstractList;
	}

	@Override
	public UserMessage adaptSettings(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		DocumentSettingsPart settingPart = documentPart.getDocumentSettingsPart();
		try {
			CTSettings settings = settingPart.getContents();
			settings.setAutoHyphenation(null);
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return msg;
	}
}
