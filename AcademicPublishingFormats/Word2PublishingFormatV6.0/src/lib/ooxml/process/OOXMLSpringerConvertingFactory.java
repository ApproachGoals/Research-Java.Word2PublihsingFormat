package lib.ooxml.process;

import java.io.File;
import java.lang.reflect.GenericArrayType;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBElement;

import org.docx4j.TraversalUtil;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.finders.ClassFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.vml.CTGroup;
import org.docx4j.vml.CTTextbox;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTSettings;
import org.docx4j.wml.CTTxbxContent;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.Pict;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STTblLayoutType;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.Text;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.PStyle;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import lib.AbstractPublishingStyle;
import lib.ooxml.author.Author;
import lib.ooxml.author.Authors;
import lib.ooxml.converter.OOXMLConverterSpringer;
import lib.ooxml.identifier.OOXMLIdentifier;
import lib.ooxml.ref.SpringerNumbering;
import lib.ooxml.style.OOXMLSpringerStyle;
import lib.ooxml.tool.OOXMLConvertTool;
import tools.LogUtils;

public class OOXMLSpringerConvertingFactory implements OOXMLConvertingInterface, Runnable {
	
	private OOXMLSpringerStyle springerStyle;
	
	private String sourceFileFormat;
	
	private boolean isAppendAbstractSpacing = false;
	
	private int titleIndex = 0;
	
	private int generatedAuthorNodeCount = 0;
	
	private File convertedFile;
	private WordprocessingMLPackage wordMLPackage;
	private MainDocumentPart documentPart;
	private StyleDefinitionsPart stylePart;
	private NumberingDefinitionsPart numberingPart;
	private AbstractPublishingStyle publishingStyle;
	
	public OOXMLSpringerConvertingFactory(File convertedFile, WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart, StyleDefinitionsPart stylePart, AbstractPublishingStyle publishingStyle) {
		// TODO Auto-generated constructor stub
		this.convertedFile = convertedFile;
		this.wordMLPackage = wordMLPackage;
		this.documentPart = documentPart;
		this.stylePart = stylePart;
		this.numberingPart = documentPart.getNumberingDefinitionsPart();
		this.publishingStyle = publishingStyle;
		AppEnvironment.getInstance().setNumberingPart(numberingPart);
	}
	
	private void appendLayoutP(MainDocumentPart documentPart, P p, int docContentIndex, PgMar pgMar){
		P newP = Context.getWmlObjectFactory().createP();
		newP.setPPr(Context.getWmlObjectFactory().createPPr());
		newP.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
		newP.getPPr().getSectPr().setPgMar(pgMar);
		
		OOXMLConvertTool.insertPAt(documentPart, newP, docContentIndex);
	}
	
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
				String s = p.toString();
				s = s.substring(s.indexOf("Abstract")+"Abstract".length()+2, s.length());
				p.getContent().clear();
				
				R newR = Context.getWmlObjectFactory().createR();
				Text text = Context.getWmlObjectFactory().createText();
				text.setValue("Abstract. ");
				OOXMLConvertTool.insertRIntoP(p, newR, 0);
				
				newR = Context.getWmlObjectFactory().createR();
				text = Context.getWmlObjectFactory().createText();
				text.setSpace("preserve");
				text.setValue(s);
				OOXMLConvertTool.insertRIntoP(p, newR, 0);
				
				p.setPPr(Context.getWmlObjectFactory().createPPr());
				p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
				p.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_ABSTRACT).getStyleId());
				
			} else {
				P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				abstractTextP.setPPr(new PPr());
				OOXMLConvertTool.removeAllRPROfP(abstractTextP);
				abstractTextP.getPPr().setPStyle(new PStyle());
				abstractTextP.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_ABSTRACT).getStyleId());
				
				R newR = Context.getWmlObjectFactory().createR();
				newR.setRPr(Context.getWmlObjectFactory().createRPr());
				newR.getRPr().setI(new BooleanDefaultTrue());
				newR.getRPr().setB(new BooleanDefaultTrue());
				Text text = Context.getWmlObjectFactory().createText();
				text.setValue("Abstract. ");
				text.setSpace("preserve");
				newR.getContent().add(text);
				OOXMLConvertTool.insertRIntoP(abstractTextP, newR, 0);
				
				OOXMLConvertTool.removePAt(documentPart, docContentIndex);
			}
			
//			Ind ind = Context.getWmlObjectFactory().createPPrBaseInd();
//			ind.setLeft(new BigInteger("567"));
//			ind.setRight(new BigInteger("567"));
//			pPr.setInd(ind);
//			
//			pPr.setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//			pPr.getSpacing().setBefore(new BigInteger("600"));
//			pPr.getSpacing().setAfter(new BigInteger("120"));
			
			appendAbstractSpacing(p);
			
			msg.setMessageCode(UserMessage.MESSAGE_ABSTRACT_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
			
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
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
				
				p.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_FIGURE).getStyleId());
				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "Figure", "Fig.", springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_FIGURECHAR).getStyleId());
				
				msg.setMessageCode(UserMessage.MESSAGE_CAPTION_ADAPT_FINISH);
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
	public UserMessage adaptFontStyleClass(File convertedFile, WordprocessingMLPackage wordMLPackage,
			MainDocumentPart documentPart, StyleDefinitionsPart stylePart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		int index = 0;
		
		while(index < springerStyle.getStyleList().size()){
			try {
				if(OOXMLConvertTool.isStyleExist(springerStyle.getStyleList().get(index), stylePart)){
					Style style = stylePart.getStyleById(springerStyle.getStyleList().get(index).getStyleId());
					int styleIndex = style!=null?OOXMLConvertTool.getStyleDefinitionIndex(style, stylePart):-1;
					if(styleIndex>=0){
//						stylePart.getContents().getStyle().remove(styleIndex);
						stylePart.getContents().getStyle().set(styleIndex, style);
					}
				} else {
//					documentPart.addStyledParagraphOfText(ieeeStyle.getStyleList().get(index).getStyleId(), "doc "+ieeeStyle.getStyleList().get(index).getName().getVal());
//					stylePart.getContents().getStyle().add(ieeeStyle.getStyleList().get(index));
					ObjectFactory factory = Context.getWmlObjectFactory ();
					Style newStyle = factory.createStyle();
					newStyle.setCustomStyle(true);
					newStyle.setBasedOn(springerStyle.getStyleList().get(index).getBasedOn());
					newStyle.setLink(springerStyle.getStyleList().get(index).getLink());
					newStyle.setName(springerStyle.getStyleList().get(index).getName());
					newStyle.setNext(springerStyle.getStyleList().get(index).getNext());
					newStyle.setPPr(springerStyle.getStyleList().get(index).getPPr());
					newStyle.setQFormat(springerStyle.getStyleList().get(index).getQFormat());
					newStyle.setRPr(springerStyle.getStyleList().get(index).getRPr());
					newStyle.setStyleId(springerStyle.getStyleList().get(index).getStyleId());
					newStyle.setTblPr(springerStyle.getStyleList().get(index).getTblPr());
					newStyle.setTcPr(springerStyle.getStyleList().get(index).getTcPr());
					newStyle.setTrPr(springerStyle.getStyleList().get(index).getTrPr());
					newStyle.setType(springerStyle.getStyleList().get(index).getType());
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
			pStyle.setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING1).getStyleId());
			pPr.setPStyle(pStyle);
			
			OOXMLConvertTool.removeHeadingNum(1, p);
			
			msg.setMessageCode(UserMessage.MESSAGE_HEADING1_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
			
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
			pStyle.setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING2).getStyleId());
			pPr.setPStyle(pStyle);
			
			OOXMLConvertTool.removeHeadingNum(2, p);
			
			msg.setMessageCode(UserMessage.MESSAGE_HEADING2_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
			
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
			if(sectionStr!=null && sectionStr.equals(OOXMLConvertTool.getPText(p))){
				PPr pPr = p.getPPr();
				if(pPr!=null){
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					pPr.getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_DEFAULTPARAGRAPH).getStyleId());
				} else {
					pPr = Context.getWmlObjectFactory().createPPr();
					p.setPPr(pPr);
				}
				ParaRPr rPr = pPr.getRPr();
				if(rPr!=null){
					pPr.setRPr(Context.getWmlObjectFactory().createParaRPr());
				} else {
					rPr = Context.getWmlObjectFactory().createParaRPr();
					pPr.setRPr(rPr);
				}
				if(rPr.getRStyle()!=null){
					
				} else {
					rPr.setRStyle(Context.getWmlObjectFactory().createRStyle());
				}
				rPr.getRStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING3CHAR).getStyleId());
				
				for(Object o : p.getContent()){
					if(o instanceof R){
						R r = (R)o;
						r.setRPr(Context.getWmlObjectFactory().createRPr());
						r.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						r.getRPr().getRStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING3CHAR).getStyleId());
						
					}
				}
				if(sectionStr.endsWith(".")){
					
				} else {
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					newR.getContent().add(text);
					text.setValue(". ");
					text.setSpace("preserve");
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					newR.getRPr().getRStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING3CHAR).getStyleId());
					p.getContent().add(newR);
				}
				
				P nextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				removePSetting(nextP.getPPr());
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
					pPr.setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					pPr.getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_DEFAULTPARAGRAPH).getStyleId());
				} else {
					pPr = Context.getWmlObjectFactory().createPPr();
					p.setPPr(pPr);
				}
				ParaRPr rPr = pPr.getRPr();
				if(rPr!=null){
					pPr.setRPr(Context.getWmlObjectFactory().createParaRPr());
				} else {
					rPr = Context.getWmlObjectFactory().createParaRPr();
					pPr.setRPr(rPr);
				}
				if(rPr.getRStyle()!=null){
					
				} else {
					rPr.setRStyle(Context.getWmlObjectFactory().createRStyle());
				}
				rPr.getRStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING4CHAR).getStyleId());
				for(Object o : p.getContent()){
					if(o instanceof R){
						R r = (R)o;
						r.setRPr(Context.getWmlObjectFactory().createRPr());
						r.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
						r.getRPr().getRStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING4CHAR).getStyleId());
						
					}
				}
				if(sectionStr.endsWith(".")){
					
				} else {
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					newR.getContent().add(text);
					text.setValue(". ");
					text.setSpace("preserve");
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					newR.getRPr().getRStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_HEADING4CHAR).getStyleId());
					p.getContent().add(newR);
				}
				
				P nextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				removePSetting(nextP.getPPr());
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

	public UserMessage adaptImages2(MainDocumentPart documentPart){
		UserMessage msg = new UserMessage();
		
		try {
			ClassFinder classFiner = new ClassFinder(Pict.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			for(Object pictObj : classFiner.results){
				Pict pict = (Pict) pictObj;
				List<Object> pictContentList = pict.getAnyAndAny();
				if(pictContentList!=null){
					for(Object o : pictContentList){
						if(o instanceof CTGroup){
							CTGroup group = (CTGroup) o;
							String styleString = group.getStyle();
							LogUtils.log(styleString);
						} else if(o instanceof JAXBElement<?>){
							if(((JAXBElement<?>) o).getValue() instanceof CTGroup){
								CTGroup group = (CTGroup) ((JAXBElement<?>) o).getValue();
								String styleString = group.getStyle();
								float factor = -1;
								int width = 345;//270;//648;
								if(styleString.indexOf(";width:")<0){
									factor = -2;
								}
								if(styleString!=null && styleString.length()>0){
									String[] temps = styleString.split(";");
									for (int i = 0; i < temps.length; i++) {
										if(temps[i].indexOf("left")==0){
											temps[i] = "left:0";
										} else if(temps[i].indexOf("margin-left:")==0){
											temps[i] = "margin-left:0pt";
										} else if(temps[i].indexOf("margin-top:")==0){
											temps[i] = "margin-top:30pt";
										} else if(temps[i].indexOf("width:")==0) {
											String[] vals = temps[i].split(":");
											if(vals[1].indexOf("pt")>0){
												String val = vals[1].substring(0, vals[1].indexOf("pt"));
												if(Float.parseFloat(val)>width){
													factor = width/Float.parseFloat(val);//-2;
													temps[i] = "width:"+width+"pt";
												} else {
													factor = width/Float.parseFloat(val);
//													LogUtils.log(width+"/"+Float.parseFloat(val)+"="+factor);
													temps[i] = "width:"+width+"pt";
												}
											} else if(vals[1].indexOf("in")>0){
												String val = vals[1].substring(0, vals[1].indexOf("in"));
												if(Float.parseFloat(val)*72>width){
													factor = width/Float.parseFloat(val);//-2;
													temps[i] = "width:"+width+"pt";
												} else {
													factor = width/(Float.parseFloat(val)*72);
													temps[i] = "width:"+width+"pt";
												}
											}
//											LogUtils.log("factor="+factor);
										} else if(temps[i].indexOf("height:")==0) {
											if(factor>0){
												String[] vals = temps[i].split(":");
												if(vals[1].indexOf("pt")>0){
													String val = vals[1].substring(0, vals[1].indexOf("pt"));
													temps[i] = "height:"+factor*Float.parseFloat(val)+"pt";
												} else if(vals[1].indexOf("in")>0){
													String val = vals[1].substring(0, vals[1].indexOf("in"));
//													LogUtils.log(factor+"*"+Float.parseFloat(val)+"*72="+factor*(Float.parseFloat(val)*72));
													temps[i] = "height:"+factor*(Float.parseFloat(val)*72)+"pt";
												}
											}
										}
									}
									styleString="";
									for (int i = 0; i < temps.length; i++) {
										boolean isEnd = (i+1>=temps.length);
										styleString+=temps[i]+(isEnd?"":";");
									}
								}
//								LogUtils.log(styleString);
								group.setStyle(styleString);
							} else {
								LogUtils.log(((JAXBElement) o).getValue().getClass().getName());
							}
						} else {
							LogUtils.log(o.getClass().getName());
						}
					}
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		try {
			ClassFinder classFiner = new ClassFinder(Inline.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			for(Object inlineObj : classFiner.results){
				Inline inline = (Inline) inlineObj;
				CTPositiveSize2D extent = inline.getExtent();
				long cx = extent.getCx();
				long cy = extent.getCy();
				if(cx>12240){
					float ratio = (float)(3*914400)/cx;
					extent.setCx(3*914400);
					long calY = (long) (cy*ratio);//(long) (cy*3/cx*914400);
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
			msg.setMessageCode(UserMessage.MESSAGE_IMAGE_ADAPT_FINISH);
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
	public UserMessage adaptImages(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		/*try {
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
//			msg.setMessageCode(UserMessage.MESSAGE_IMAGE_ADAPT_FINISH);
//			LogUtils.log(msg.getMessage());
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}*/
		
		adaptImages2(documentPart);
		
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
					|| OOXMLConvertTool.isPTextStartWithString(p, "Keywords.", true))){
				if(p.toString().length()>"Keywords".length()+2){
					String s = p.toString();
					s = s.substring("Keywords".length()+2, s.length());
					p.getContent().clear();
					
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue("Keywords: ");
					text.setSpace("preserve");
					OOXMLConvertTool.insertRIntoP(p, newR, 0);
					
					newR = Context.getWmlObjectFactory().createR();
					text = Context.getWmlObjectFactory().createText();
					text.setValue(s);
					OOXMLConvertTool.insertRIntoP(p, newR, 0);
					
					p.setPPr(Context.getWmlObjectFactory().createPPr());
					p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					p.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_KEYWORDHEAD).getStyleId());

				} else {
					P keywordTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
					keywordTextP.setPPr(new PPr());
					OOXMLConvertTool.removeAllRPROfP(keywordTextP);
					keywordTextP.getPPr().setPStyle(new PStyle());
					keywordTextP.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_KEYWORDHEAD).getStyleId());
					
					R newR = Context.getWmlObjectFactory().createR();
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setI(new BooleanDefaultTrue());
					newR.getRPr().setB(new BooleanDefaultTrue());
					Text text = Context.getWmlObjectFactory().createText();
					text.setValue("Keywords: ");
					text.setSpace("preserve");
					newR.getContent().add(text);
					OOXMLConvertTool.insertRIntoP(keywordTextP, newR, 0);
					
					OOXMLConvertTool.removePAt(documentPart, docContentIndex);
				}
				
			} else {
				
			}
			
//			Ind ind = Context.getWmlObjectFactory().createPPrBaseInd();
//			ind.setLeft(new BigInteger("567"));
//			ind.setRight(new BigInteger("567"));
//			pPr.setInd(ind);
//			
//			pPr.setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//			pPr.getSpacing().setBefore(new BigInteger("600"));
//			pPr.getSpacing().setAfter(new BigInteger("120"));
			
			appendAbstractSpacing(p);
			isAppendAbstractSpacing = false;
			
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
				pPr.getPStyle().setVal(springerStyle.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHY).getStyleId());
				pPr.setInd(new Ind());
				pPr.getInd().setLeft(new BigInteger("360"));
				pPr.getInd().setHanging(new BigInteger("360"));
				
				msg.setMessageCode(UserMessage.MESSAGE_REFERENCE_ADPAT_FINISH);
				if(AppEnvironment.getInstance().isActiveTestMode()){
					LogUtils.log(msg.getMessage());
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
				pPr.getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_REFERENCEHEAD).getStyleId());
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
		SpringerNumbering.adaptNumberingFile(wordMLPackage, documentPart);
		msg.setMessageCode(UserMessage.MESSAGE_NUMBERFILE_ADAPT_FINISH);
		LogUtils.log(msg.getMessage());
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
					OOXMLConvertTool.fixPageLayout(bodySectPr, "1", "720", false);
					/*
					 * fix: default header need to remove
					 */
					if(bodySectPr.getEGHdrFtrReferences()!=null){
						bodySectPr.getEGHdrFtrReferences().clear();
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
			pStyle.setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_TITLE).getStyleId());
			pPr.setPStyle(pStyle);
			
			msg.setMessageCode(UserMessage.MESSAGE_TITLE_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
			
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
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
			int pageWidth = 0;
			try {
				pageWidth = AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT).getPPr().getSectPr().getPgSz().getW().intValue();
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
			while(index<classFinder.results.size()){
				Tbl tbl = (Tbl) classFinder.results.get(index);
				if(tbl.getTblPr()!=null){
//					tbl.getTblPr().getTblCaption();
					if(tbl.getTblPr().getTblW()!=null){
							//&& tbl.getTblPr().getTblW().getType()!=null
							//&& !tbl.getTblPr().getTblW().getType().equals("auto")
							//&& tbl.getTblPr().getTblW().getW()!=null
							//&& tbl.getTblPr().getTblW().getW().intValue()>pageWidth
							//&& tbl.getTblGrid()!=null
							//&& tbl.getTblGrid().getGridCol()!=null
							//&& tbl.getTblGrid().getGridCol().size()<6){
//						System.out.println("something in 1");
						tbl.getTblPr().getTblW().setW(new BigInteger(""+pageWidth));
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
						int calDiff = pageWidth-calTblW;
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
						try {
							tbl.getTblPr().getTblW().setW(new BigInteger(""+pageWidth));
						} catch (Exception e) {
							// TODO: handle exception
							LogUtils.log(e.getMessage());
						}
						
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
				p.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_TABLEHEADER).getStyleId());
//				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setRPr(null);
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "TABLE", "Table", springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_TABLEHEADERCHAR).getStyleId());
				
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
			newP.getPPr().getSectPr().getCols().setNum(BigInteger.valueOf(cols));
			newP.getPPr().getSectPr().getCols().setSpace(new BigInteger("720"));
			newP.getPPr().getSectPr().setType(Context.getWmlObjectFactory().createSectPrType());
			newP.getPPr().getSectPr().getType().setVal("continuous");
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
			springerStyle = new OOXMLSpringerStyle();
			Authors authors = new Authors();
			
			int docContentIndex = 0;
			int authorCount = 0;
			boolean isAuthorbeginn = false;
			boolean isKeywordBegin = false;
			boolean isAbstractBegin = false;
			int authorInfoLine = 0;
			int authorNodeCount = 0;
			
			List<Object> docContentList = documentPart.getContent();
			
			adaptSettings(documentPart);
			adaptSectPrs(documentPart);
			
			removeUselessPageBreaks(wordMLPackage, documentPart);
			
			ClassFinder classFiner = new ClassFinder(CTTextbox.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			for(Object pObj : classFiner.results){
				CTTextbox tb = (CTTextbox)pObj;
				if(tb.getTxbxContent()!=null){
					CTTxbxContent txbxContent = tb.getTxbxContent();
					for(Object o : txbxContent.getContent()){
						if(o instanceof P){
							if(OOXMLConvertTool.isCaption((P)o, documentPart, stylePart)){
								adaptCaption(documentPart, (P)o);
							}
						}
					}
				}
			}
			
			while(docContentIndex<docContentList.size()){
				Object o = docContentList.get(docContentIndex);
				if(o instanceof P) {
					P p = (P) o;
					if(OOXMLConvertTool.isIncludeSectPr(p)){
						OOXMLConvertTool.removeSectPr(p);
						if(OOXMLConvertTool.isEmpty(p)){
							docContentIndex = OOXMLConvertTool.removeNode(docContentList, docContentIndex);
						}
					} else if(OOXMLConvertTool.isTitle(p, stylePart)){
						titleIndex = docContentIndex;
						adaptTitle(documentPart, p);
//						appendLayoutPara(documentPart, p, docContentIndex, 1);
//						docContentIndex++;
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
						adaptAuthors(documentPart, docContentIndex, authorNodeCount, authors);
						docContentIndex = docContentIndex-authorNodeCount+generatedAuthorNodeCount;
						adaptAbstract(documentPart, p, docContentIndex);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isKeywords(p, stylePart)) {
						adaptKeyWord(documentPart, p, docContentIndex);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isSectionHeader1(p, stylePart, documentPart)){
						adaptHeading1(p);
						saveChange(convertedFile, wordMLPackage);
					} else if(OOXMLConvertTool.isSectionHeader2(p, stylePart, documentPart)) {
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
//					} else if(OOXMLConvertTool.isEmpty(p)){
//						docContentIndex = OOXMLConvertTool.removeNode(docContentList, docContentIndex);
					} else {
						if(isAppendAbstractSpacing){
							appendAbstractSpacing(p);
						}
						if(isAuthorbeginn){
							authorInfoLine++;
							authors.getAuthors().get(authorCount-1).setInfo(authorInfoLine, p.toString());
							authorNodeCount++;
						}
						
					}
					
				} else {
					
				}
				docContentIndex++;
			}
			adaptFootNote(convertedFile, wordMLPackage);
			adaptImages(documentPart);
			adaptFontStyleClass(convertedFile, wordMLPackage, documentPart, stylePart);
			adaptPageLayout(documentPart, wordMLPackage);
			adaptNumberFile(documentPart);
			adaptTable(documentPart);
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
		UserMessage msg = new UserMessage();
		try {
			ClassFinder classFiner = new ClassFinder(P.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			for(int index =0; index<classFiner.results.size(); index++){
				Object o = classFiner.results.get(index);
				P p = (P)o;
				if(p.getPPr()!=null && p.getPPr().getSectPr()!=null){
					p.getPPr().setSectPr(null);
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage saveChange(File convertedFile, WordprocessingMLPackage wordMLPackage) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		try {
			wordMLPackage.save(convertedFile);
			msg.setMessageDetails("The Changes have been saved.");
			LogUtils.log(msg.getMessage());
		} catch (Exception e) {
			// TODO: handle exception
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		return msg;
	}

	private void adaptSectPrs(MainDocumentPart documentPart){
		try {
			ClassFinder classFiner = new ClassFinder(P.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
			int i=1;
			for(Object o : classFiner.results){
				P p = (P)o;
				if(p.getPPr()!=null && p.getPPr().getSectPr()!=null){
					p.getPPr().setSectPr(null);
//					SectPr sectPr = p.getPPr().getSectPr();
////					OOXMLConvertTool.adaptPageLayout(sectPr, null, null, "10773", "14742", "93", "754", "2126", "737", "1191", "907", "1202", "0");
//					OOXMLConvertTool.adaptPageLayout(sectPr, "1", "720", "10773", "14742", "93", null, null, "737", "1191", "907", "1202", "0");
				}
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
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
			newP.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_AUTHORS).getStyleId());
			R r = Context.getWmlObjectFactory().createR();
			Text text = Context.getWmlObjectFactory().createText();
			text.setValue(authors.getNameString());
			
			r.getContent().add(text);
			newP.getContent().add(r);
			documentPart.getContent().add(index, newP);
			int i=0;
			for(Author author: authors.getAuthors()){
				if(author.getInfo()!=null && author.getInfo().length()>0){
					newP = Context.getWmlObjectFactory().createP();
					newP.setPPr(Context.getWmlObjectFactory().createPPr());
					newP.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
					if(i+1>=authors.getAuthors().size()){
						newP.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_EMAIL).getStyleId());
					} else{
						newP.getPPr().getPStyle().setVal(springerStyle.getStyleMap().get(OOXMLSpringerStyle.STYLENAME_AFFILIATIONS).getStyleId());
					}
					r = Context.getWmlObjectFactory().createR();
					text = Context.getWmlObjectFactory().createText();
					text.setValue(author.getInfo());
					
					r.getContent().add(text);
					newP.getContent().add(r);
					documentPart.getContent().add(index+(++i), newP);
					
				}
			}
			generatedAuthorNodeCount = i;
			msg.setMessageCode(UserMessage.MESSAGE_AUTHORS_ADAPT_FINISH);
			LogUtils.log(msg.getMessage());
		}
		return msg;
	}

	@Override
	public UserMessage adaptSettings(MainDocumentPart documentPart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		DocumentSettingsPart settingPart = documentPart.getDocumentSettingsPart();
		try {
			CTSettings settings = settingPart.getContents();
			settings.setAutoHyphenation(new BooleanDefaultTrue());
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return msg;
	}

	private void appendAbstractSpacing(P p){
		PPr pPr = OOXMLConvertTool.getPPr(p, true, true);
		
		Ind ind = Context.getWmlObjectFactory().createPPrBaseInd();
		ind.setLeft(new BigInteger("567"));
		ind.setRight(new BigInteger("567"));
		pPr.setInd(ind);
		
		pPr.setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
		pPr.getSpacing().setBefore(new BigInteger("600"));
		pPr.getSpacing().setAfter(new BigInteger("120"));
		
		isAppendAbstractSpacing = true;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		convert2FormatNew(convertedFile, wordMLPackage, documentPart, stylePart);
	}
	
	private UserMessage convert2FormatNew(File convertedFile, WordprocessingMLPackage wordMLPackage,
			MainDocumentPart documentPart, StyleDefinitionsPart stylePart){
		UserMessage msg = new UserMessage();
		try {
			NumberingDefinitionsPart numberingPart = documentPart.getNumberingDefinitionsPart();
			List<Object> docContentList = documentPart.getContent();
			springerStyle = (OOXMLSpringerStyle) publishingStyle;//new OOXMLSpringerStyle();
			if(AppEnvironment.getInstance().getNewStyleFile()!=null){
				springerStyle.loadStyleFromFile(AppEnvironment.getInstance().getNewStyleFile());
			} else {
				springerStyle.setStyleMap(AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.SPRINGER).getStyleMap());	
			}
			
			OOXMLConverterSpringer converter = new OOXMLConverterSpringer();
			OOXMLIdentifier identifier = new OOXMLIdentifier();
			identifier.init(stylePart);
			
			converter.removeDefaultHeader(documentPart);
			
			int docContentIndex = 0;
			boolean isHeadAdapted = false;
			while(docContentIndex<docContentList.size()){
				Object o = docContentList.get(docContentIndex);
				if(o instanceof P) {
					P p = (P) o;
					String pText = OOXMLConvertTool.getPText(p);
//					if(pText.indexOf("Shafiq")>=0){
//						System.out.println(pText);
//					}
					int headingLvl = 0;
					if(AppEnvironment.getInstance().getGui().getIsActiveTestMode().isSelected()){
						LogUtils.log(OOXMLConvertTool.getPText(p));
					}
					if(identifier.isKeywordIdentified()){
						if(OOXMLConvertTool.isIncludeSectPr(p)){
							OOXMLConvertTool.removeSectPr(p);
							if(OOXMLConvertTool.isEmpty(p)){
								OOXMLConvertTool.removeNode(docContentList, docContentIndex);
								continue;
							}
						}
						headingLvl = identifier.getHeadingLvl(p, numberingPart, stylePart);
					} else if(identifier.isAbstractIdentified()) {
						if(OOXMLConvertTool.isIncludeSectPr(p)){
							int col = 1;
							if(p.getPPr().getSectPr().getCols()!=null){
								col = p.getPPr().getSectPr().getCols().getNum().intValue();
							}
							OOXMLConvertTool.changeSectPr(p, col, springerStyle.getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT));
						}
					}
					if(identifier.isTitle(p, stylePart)){
						converter.adaptTitle(documentPart, p);
						converter.appendLayoutPara(documentPart, p, docContentIndex, 1);
//						docContentIndex++;
//					} else if(identifier.isSubtitle(p, stylePart)) {
//						converter.adaptSubtitle(documentPart, p);
//						converter.appendLayoutPara(documentPart, Context.getWmlObjectFactory().createP(), docContentIndex, 1);
					} else if(identifier.isAbstractHeader(p, stylePart)){
						converter.appendLayoutPara(documentPart, p, docContentIndex-1, 1);
						docContentIndex++;
						converter.adaptAbstract(documentPart, p, docContentIndex);
//						docContentIndex++;
						isHeadAdapted = true;
					} else if(identifier.isKeywordHeader(p, stylePart)){
						converter.adaptKeyWord(documentPart, p, docContentIndex);
//						docContentIndex++;
						isHeadAdapted = true;
					} else if(identifier.isAcknowledgmentHeader(p, numberingPart, stylePart)){
						converter.adaptAcknowledgment(documentPart, p, docContentIndex);
//						docContentIndex++;
						isHeadAdapted = true;
					} else if(identifier.isReferenceHeader(p, numberingPart, stylePart)) {
						converter.adaptLiteratureHeader(documentPart, p);
					} else if(headingLvl>0){
						switch (headingLvl) {
						case 1:
							converter.adaptHeading1(p);
							isHeadAdapted = true;
							break;
						case 2:
							converter.adaptHeading2(p);
							isHeadAdapted = true;
							break;
						case 3:
							converter.adaptHeading3(documentPart, p, docContentIndex);
							isHeadAdapted = true;
							break;
						case 4:
							converter.adaptHeading4(documentPart, p, docContentIndex);
							isHeadAdapted = true;
							break;
						default:
							break;
						}
					} else if(identifier.isFigureCaption(p, stylePart, numberingPart)){
						converter.adaptFigureCaption(documentPart, p);
					} else if(identifier.isTableCaption(p, stylePart, numberingPart)){
						converter.adaptTableCaption(documentPart, p);
					} else if(identifier.isAppendix(p, stylePart)){
						converter.appendLayoutPara(documentPart, p, docContentIndex-1, 2);
						docContentIndex++;
						converter.adaptAppendixHeader(documentPart, p, docContentIndex);
					} else if(identifier.isAppendixHeaderIdentified() && docContentIndex>documentPart.getContent().size()-2){
						converter.appendLayoutPara(documentPart, p, docContentIndex, 1);
//					} else if(identifier.isReferencesHeaderIdentified()){
//						adaptLiterature(documentPart, p);
					} else {
						if(p.getPPr()!=null && p.getPPr().getFramePr()!=null){
							OOXMLConvertTool.fixFramePr(p);
						}
						if(isHeadAdapted){
							OOXMLConvertTool.fixFirstLineEmptySpaces(p, true, null);
						}
						isHeadAdapted = false;
						if(identifier.isAcknowledgeIdentified()){
							OOXMLConvertTool.removeSectPr(p);
						} else if(identifier.isTitleIdentified() && !identifier.isAbstractIdentified()){
							OOXMLConvertTool.replaceSectPr(p, springerStyle.getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT));//if sectPr exists, set the page size and margin
							OOXMLConvertTool.setAlignmentCenter(p);// fix p alignment to center
							OOXMLConvertTool.removeInd(p);
						}
						
						if(identifier.isReferencesHeaderIdentified() && !identifier.isAppendixHeaderIdentified()){
							adaptLiterature(documentPart, p);
						} else {
							OOXMLConvertTool.setPStyle(p, springerStyle.getStyleMap().get(AcademicFormatStructureDefinition.NORMAL).getStyleId());
						}
					}
				} else if(o instanceof SdtBlock){	// generated bibliography
					docContentIndex += OOXMLConvertTool.adaptGeneratedReferences((SdtBlock)o, springerStyle.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER).getStyleId(), springerStyle.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHY).getStyleId(), docContentIndex, docContentList) - 1;
				} else if(o instanceof Tbl){
					Tbl tbl = (Tbl)o;
				}
				docContentIndex++;
				AppEnvironment.getInstance().getGui().getProgressBar().setValue(docContentIndex*100/(6+docContentList.size()));
			}
			
			converter.adaptFootNote(convertedFile, wordMLPackage);
			AppEnvironment.getInstance().getGui().getProgressBar().setValue((++docContentIndex)*100/(6+docContentList.size()));
			adaptImages(documentPart);
			AppEnvironment.getInstance().getGui().getProgressBar().setValue((++docContentIndex)*100/(6+docContentList.size()));
			converter.adaptFontStyleClass(convertedFile, wordMLPackage, documentPart, stylePart);
			AppEnvironment.getInstance().getGui().getProgressBar().setValue((++docContentIndex)*100/(6+docContentList.size()));
			converter.adaptPageLayout(documentPart, wordMLPackage);
			AppEnvironment.getInstance().getGui().getProgressBar().setValue((++docContentIndex)*100/(6+docContentList.size()));
			adaptNumberFile(documentPart);
			AppEnvironment.getInstance().getGui().getProgressBar().setValue((++docContentIndex)*100/(6+docContentList.size()));
			adaptTable(documentPart);
			AppEnvironment.getInstance().getGui().getProgressBar().setValue((++docContentIndex)*100/(6+docContentList.size()));
			
			saveChange(convertedFile, wordMLPackage);
			
		} catch (Exception e) {
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		msg.setMessageCode(UserMessage.MESSAGE_GENERAL_OK);
		LogUtils.log(msg.getMessage());
		
		return msg;
	}
}
