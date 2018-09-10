package lib.ooxml.process;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;
import java.util.Iterator;
import java.util.List;

import javax.xml.bind.JAXBElement;

import org.apache.commons.io.IOUtils;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.dml.CTPositiveSize2D;
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
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.sharedtypes.STOnOff;
import org.docx4j.vml.CTGroup;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTRel;
import org.docx4j.wml.CTSettings;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.P.Hyperlink;
import org.docx4j.wml.PPr;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.Pict;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STTblLayoutType;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.PPrBase.PStyle;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import lib.AbstractPublishingStyle;
import lib.ooxml.author.Author;
import lib.ooxml.author.Authors;
import lib.ooxml.converter.OOXMLConverterElsevier;
import lib.ooxml.identifier.OOXMLIdentifier;
import lib.ooxml.ref.ElsevierNumbering;
import lib.ooxml.style.OOXMLElsevierStyle;
import lib.ooxml.tool.OOXMLConvertTool;
import lib.ooxml.tool.OOXMLElsevierHeaderFooterProvider;
import tools.LogUtils;

public class OOXMLElsevier1ConvertingFactory implements OOXMLConvertingInterface, Runnable {

	private OOXMLElsevierStyle elsevierStyle;
	
	private String sourceFileFormat;
	
	private int titleIndex = 0;
	
	private File convertedFile;
	private WordprocessingMLPackage wordMLPackage;
	private MainDocumentPart documentPart;
	private StyleDefinitionsPart stylePart;
	private NumberingDefinitionsPart numberingPart;
	private AbstractPublishingStyle publishingStyle;
	
	public OOXMLElsevier1ConvertingFactory(File convertedFile, WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart, StyleDefinitionsPart stylePart, AbstractPublishingStyle publishingStyle){
		this.convertedFile = convertedFile;
		this.wordMLPackage = wordMLPackage;
		this.documentPart = documentPart;
		this.stylePart = stylePart;
		this.numberingPart = documentPart.getNumberingDefinitionsPart();
		this.publishingStyle = publishingStyle;
		AppEnvironment.getInstance().setNumberingPart(numberingPart);
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
				R newR = new R();
				Text text = new Text();
				text.setValue("Abstract");
				newR.getContent().add(text);
				P newP = new P();
				newP.getContent().add(newR);
				newP.setPPr(new PPr());
				newP.getPPr().setPStyle(new PStyle());
				newP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_ABSTRACTHEADER).getStyleId());
				
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
				pPr.getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_ABSTRACT).getStyleId());
				p.setPPr(pPr);
				
			} else {
				PStyle pStyle = OOXMLConvertTool.getPStyle(pPr, true);
				pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_ABSTRACTHEADER).getStyleId());
				pPr.setPStyle(pStyle);
				
				P abstractTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				abstractTextP.setPPr(new PPr());
				OOXMLConvertTool.removeAllRPROfP(abstractTextP);
				abstractTextP.getPPr().setPStyle(new PStyle());
				abstractTextP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_ABSTRACT).getStyleId());
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
				p.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_CAPTION).getStyleId());
				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "Figure", "Fig.", null);
				
				msg.setMessageCode(UserMessage.MESSAGE_CAPTION_ADAPT_FINISH);
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
	public UserMessage adaptFontStyleClass(File convertedFile, WordprocessingMLPackage wordMLPackage,
			MainDocumentPart documentPart, StyleDefinitionsPart stylePart) {
		// TODO Auto-generated method stub
		UserMessage msg = new UserMessage();
		
		int index = 0;
		
		while(index < elsevierStyle.getStyleList().size()){
			try {
				if(OOXMLConvertTool.isStyleExist(elsevierStyle.getStyleList().get(index), stylePart)){
					Style style = stylePart.getStyleById(elsevierStyle.getStyleList().get(index).getStyleId());
//					int styleIndex = style!=null?OOXMLConvertTool.getStyleDefinitionIndex(style, stylePart):-1;
//					if(styleIndex>=0){
////						stylePart.getContents().getStyle().remove(styleIndex);
//						stylePart.getContents().getStyle().set(styleIndex, style);
//					}
					style = elsevierStyle.getStyleList().get(index);
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
			} catch (Exception e) {
				// TODO Auto-generated catch block
				LogUtils.log("ERROR while adapting style." + e.getMessage());
			}
			index++;
		}
		
		saveChange(convertedFile, wordMLPackage);
		if(AppEnvironment.getInstance().isActiveTestMode()){
			LogUtils.log("font class will be modified.");
		}
		
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
			pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_HEADING1).getStyleId());
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
			pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_HEADING2).getStyleId());
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
				rPr.getRStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_HEADING3).getStyleId());
				
				if(sectionStr.endsWith(".")){
					
				} else {
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					newR.getContent().add(text);
					text.setValue(". ");
					text.setSpace("preserve");
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					newR.getRPr().getRStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_HEADING3).getStyleId());
					p.getContent().add(newR);
				}

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
				rPr.getRStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_HEADING4).getStyleId());
				
				if(sectionStr.endsWith(".")){
					
				} else {
					R newR = Context.getWmlObjectFactory().createR();
					Text text = Context.getWmlObjectFactory().createText();
					newR.getContent().add(text);
					text.setValue(". ");
					text.setSpace("preserve");
					newR.setRPr(Context.getWmlObjectFactory().createRPr());
					newR.getRPr().setRStyle(Context.getWmlObjectFactory().createRStyle());
					newR.getRPr().getRStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_HEADING4).getStyleId());
					p.getContent().add(newR);
				}
				
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
								int width = 495;//270;//648;
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
		}
		*/
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
				p.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_KEYWORDHEAD).getStyleId());
				
			} else {
				P keywordTextP = OOXMLConvertTool.getNextP(documentPart, docContentIndex);
				keywordTextP.setPPr(new PPr());
				OOXMLConvertTool.removeAllRPROfP(keywordTextP);
				keywordTextP.getPPr().setPStyle(new PStyle());
				keywordTextP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_KEYWORDHEAD).getStyleId());
				
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
				pPr.getPStyle().setVal(elsevierStyle.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHY).getStyleId());
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
				pPr.getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_REFERENCEHEAD).getStyleId());
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
		ElsevierNumbering.adaptNumberingFile(wordMLPackage, documentPart);
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
//					OOXMLConvertTool.fixPageLayout(bodySectPr, "1", "720", false);
					OOXMLConvertTool.adaptPageLayout(bodySectPr, "1", "720", "10773", "14742", "93", "754", "2126", "737", "1191", "907", "1202", "0");
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
			pStyle.setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_TITLE).getStyleId());
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
				p.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_CAPTION).getStyleId());
//				p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
				p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
				p.getPPr().getJc().setVal(JcEnumeration.CENTER);
				OOXMLConvertTool.removeAllRPROfP(p);
				
				OOXMLConvertTool.adaptFirstWord2Special(p, "TABLE", "Table", null);
				
				msg.setMessageCode(UserMessage.MESSAGE_TABLECAPTION_ADAPT_FINISH);
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
			newP.getPPr().getSectPr().setTitlePg(new BooleanDefaultTrue());
			OOXMLConvertTool.adaptPageLayout(newP.getPPr().getSectPr(), String.valueOf(cols), "720", "10773", "14742", "93", null, null, null, null, null, null, null);
			
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
			elsevierStyle = new OOXMLElsevierStyle();
			
			Authors authors = new Authors();
			
			adaptPageLayout(documentPart, wordMLPackage);
			adaptSectPrs(documentPart);
			
			int docContentIndex = 0;
			int authorCount = 0;
			boolean isAuthorbeginn = false;
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
//						appendLayoutPara(documentPart, p, docContentIndex);
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
//						adaptlastNodeLayout(documentPart, docContentIndex);
//						saveChange(convertedFile, wordMLPackage);
						adaptAuthors(documentPart, docContentIndex, authorNodeCount, authors);
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
					} else {
						removePSetting(p.getPPr());
						if(isAuthorbeginn){
							authorInfoLine++;
							authors.getAuthors().get(authorCount-1).setInfo(authorInfoLine, p.toString());
							authorNodeCount++;
						}
					}
				} else if(o instanceof Tbl){
					Tbl tbl = (Tbl)o;
					
				}
				docContentIndex++;
			}
			
			adaptFootNote(convertedFile, wordMLPackage);
			adaptImages(documentPart);
			adaptFontStyleClass(convertedFile, wordMLPackage, documentPart, stylePart);
						
			adaptNumberFile(documentPart);
			adaptTable(documentPart);
			
			adaptFormatHeader(wordMLPackage);		
			
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
//			PartName partName = new PartName("/word/header/header2.xml");
			PartName partName = new PartName("/word/header2.xml");
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
//			sectionProperties.setTitlePg(new BooleanDefaultTrue());
			
			HeaderReference headerReference = factory.createHeaderReference();
			headerReference.setId(relationship.getId());
			headerReference.setType(HdrFtrRef.DEFAULT);
			sectionProperties.getEGHdrFtrReferences().add(headerReference);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
	}
	
	private void createHeaderPart(WordprocessingMLPackage wordMLPackage) {
		try {
			ObjectFactory factory = Context.getWmlObjectFactory();
//			PartName partName = new PartName("/word/header/header1.xml");
			PartName partName = new PartName("/word/header1.xml");
			HeaderPart headerPart = new HeaderPart(partName);
			headerPart.setPackage(wordMLPackage);

//			Hdr header = getHdr(wordMLPackage, headerPart);
			Hdr header = factory.createHdr();
			
//			P p = Context.getWmlObjectFactory().createP();
//			R r = Context.getWmlObjectFactory().createR();
//			Text text = Context.getWmlObjectFactory().createText();
//			p.getContent().add(r);
//			p.setPPr(Context.getWmlObjectFactory().createPPr());
//			p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
//			p.getPPr().getJc().setVal(JcEnumeration.CENTER);
//			r.getContent().add(text);
//			r.setRPr(Context.getWmlObjectFactory().createRPr());
//			r.getRPr().setI(new BooleanDefaultTrue());
//			r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
//			r.getRPr().getSz().setVal(new BigInteger("16"));
//			r.getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
//			r.getRPr().getRFonts().setAscii("Times New Roman");
//			text.setValue("Author name / Procedia Computer Science 000 (2017) 000–111");
//			
//			header.getContent().add(p);
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
			sectionProperties.getTitlePg().setVal(true);
			
//			HeaderReference headerReference = factory.createHeaderReference();
//			headerReference.setId(relationship.getId());
//			headerReference.setType(HdrFtrRef.DEFAULT);
//			sectionProperties.getEGHdrFtrReferences().add(headerReference);
			HeaderReference firstPageHeaderRef = factory.createHeaderReference();
			firstPageHeaderRef.setId(relationship.getId());
			firstPageHeaderRef.setType(HdrFtrRef.FIRST);
			sectionProperties.getEGHdrFtrReferences().add(firstPageHeaderRef);
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
			tbl.setTblPr(Context.getWmlObjectFactory().createTblPr());
			tbl.getTblPr().setTblW(Context.getWmlObjectFactory().createTblWidth());
			tbl.getTblPr().getTblW().setW(BigInteger.ZERO);
			tbl.getTblPr().getTblW().setType("auto");
			tbl.getTblPr().setTblInd(Context.getWmlObjectFactory().createTblWidth());
			tbl.getTblPr().getTblInd().setW(BigInteger.valueOf(108));
			tbl.getTblPr().getTblInd().setType("dxa");
			tbl.getTblPr().setTblLayout(Context.getWmlObjectFactory().createCTTblLayoutType());
			tbl.getTblPr().getTblLayout().setType(STTblLayoutType.FIXED);
			tbl.getTblPr().setTblLook(Context.getWmlObjectFactory().createCTTblLook());
			tbl.getTblPr().getTblLook().setVal("0000");
			tbl.getTblPr().getTblLook().setFirstColumn(STOnOff.ZERO);
			tbl.getTblPr().getTblLook().setLastColumn(STOnOff.ZERO);
			tbl.getTblPr().getTblLook().setFirstRow(STOnOff.ZERO);
			tbl.getTblPr().getTblLook().setLastRow(STOnOff.ZERO);
			tbl.getTblPr().getTblLook().setNoHBand(STOnOff.ZERO);
			tbl.getTblPr().getTblLook().setNoVBand(STOnOff.ZERO);
			
			tbl.setTblGrid(Context.getWmlObjectFactory().createTblGrid());
			
//			LogUtils.log("writableWidthTwips: "+writableWidthTwips);
			
			int[] colWidth = {1200, writableWidthTwips-1200-1920, 1920}; // cm *20*72/2.54
			
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
								
								InputStream in = this.getClass().getResourceAsStream("/sources/elsevier/elsevier_header_image1.png");
								imageFile = File.createTempFile("img1", ".png");
								FileOutputStream out = null;
								try {
									out = new FileOutputStream(imageFile);
									IOUtils.copy(in, out);
						        } finally {
						        	if(in!=null){
						        		in.close();
						        	}
						        	if(out!=null){
						        		out.close();
						        	}
								}
								
//								imageURL = getClass().getResource("/sources/elsevier/elsevier_header_image1.png");
//								imageFile = new File(imageURL.toURI());
								// 
								bytes = convertImageToByteArray(imageFile);
								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Elsevier", "Elsevier logo", 604800, 669600);
								p.setPPr(Context.getWmlObjectFactory().createPPr());
								p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
								p.getPPr().getPStyle().setVal("HeaderStyle");
								tc.getContent().add(p);
								
								imageFile.delete();
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
//								p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
//								p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//								p.getPPr().getSpacing().setLine(new BigInteger("200"));
//								p.getPPr().getSpacing().setLineRule(STLineSpacingRule.AT_LEAST);
								
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
//								p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
//								p.getPPr().getRPr().setRFonts(Context.getWmlObjectFactory().createRFonts());
//								p.getPPr().getRPr().getRFonts().setAscii("Arial");
//								p.getPPr().getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
//								p.getPPr().getRPr().getSz().setVal(new BigInteger("33"));
//								p.getPPr().setSpacing(Context.getWmlObjectFactory().createPPrBaseSpacing());
//								p.getPPr().getSpacing().setLine(new BigInteger("300"));
//								p.getPPr().getSpacing().setLineRule(STLineSpacingRule.AT_LEAST);
								
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
								
//								imageURL = getClass().getResource("/sources/elsevier/elsevier_header_cs_image3.png");
//								imageFile = new File(imageURL.toURI());
								
								InputStream in2 = this.getClass().getResourceAsStream("/sources/elsevier/elsevier_header_cs_image3.png");
								imageFile = File.createTempFile("img3", ".png");
								FileOutputStream out2 = null;
								try {
									out2 = new FileOutputStream(imageFile);
									IOUtils.copy(in2, out2);
						        } finally {
						        	if(in2!=null){
						        		in2.close();
						        	}
						        	if(out2!=null){
						        		out2.close();
						        	}
								}
								
								bytes = convertImageToByteArray(imageFile);
								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Procedia", "Procedia logo", 1137600, 741600);
								p.setPPr(Context.getWmlObjectFactory().createPPr());
								p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
								p.getPPr().getPStyle().setVal("HeaderStyle");
								tc.getContent().add(p);
								imageFile.delete();
								break;
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
	
	public static Hyperlink createHyperlink(MainDocumentPart documentPart, String url) {
		if(url!=null){
			try {
				// We need to add a relationship to word/_rels/document.xml.rels
				// but since its external, we don't use the
				// usual wordMLPackage.getMainDocumentPart().addTargetPart
				// mechanism
				org.docx4j.relationships.ObjectFactory factory = new org.docx4j.relationships.ObjectFactory();

				org.docx4j.relationships.Relationship rel = factory.createRelationship();
				rel.setType(Namespaces.HYPERLINK);
				rel.setTarget(url);
				rel.setTargetMode("External");

				documentPart.getRelationshipsPart().addRelationship(rel);

				// addRelationship sets the rel's @Id

				String hpl = "<w:hyperlink r:id=\"" + rel.getId()
						+ "\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
						+ "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" >" + "<w:r>"
						+ "<w:rPr>" + "<w:rStyle w:val=\"Hyperlink\" />" + 
						"</w:rPr>" + "<w:t>"+url+"</w:t>" + "</w:r>" + "</w:hyperlink>";

				// return (Hyperlink)XmlUtils.unmarshalString(hpl, Context.jc,
				// P.Hyperlink.class);
				return (Hyperlink) XmlUtils.unmarshalString(hpl);

			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				return null;
			}
		} else {
			return null;
		}
		

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
	
	private P addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes, HeaderPart part, String fileNameHint, String altText) {
		P paragraph = null;
		
		addImageToPackage(wordMLPackage, bytes, part, fileNameHint, altText, 0, 0);
				
		return paragraph;
	}
	private int docPrId = 1;
	private P addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes, HeaderPart part, String fileNameHint, String altText, long cx, long cy) {
		P paragraph = null;
		try {
			BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, part, bytes);

//			int docPrId = 1;
			int cNvPrId = 2;
			Inline inline = imagePart.createImageInline(fileNameHint, altText, docPrId++, cNvPrId, false);
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
	
	private void adaptlastNodeLayout(MainDocumentPart documentPart, int docContentIndex){
		Object o = documentPart.getContent().get(--docContentIndex);
		if(docContentIndex<=1){
			// finish
		} else if(o instanceof P){
			P p = (P)o;
			if(p.getPPr()!=null && p.getPPr().getSectPr()!=null){
				p.getPPr().getSectPr();
				OOXMLConvertTool.adaptPageLayout(p.getPPr().getSectPr(), null, null, "10773", "14742", "93", "754", "2126", "737", "1191", "907", "1202", "0");
			}
		} else {
			adaptlastNodeLayout(documentPart, docContentIndex);
		}
	}
	
	private void adaptSectPrs(MainDocumentPart documentPart){
		try {
			ClassFinder classFiner = new ClassFinder(P.class);
			new TraversalUtil(documentPart.getContent(), classFiner);
//			int i=1;
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
			newP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_AUTHORS).getStyleId());
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
					newP.getPPr().getPStyle().setVal(elsevierStyle.getStyleMap().get(OOXMLElsevierStyle.STYLENAME_AFFILIATIONS).getStyleId());
					r = Context.getWmlObjectFactory().createR();
					text = Context.getWmlObjectFactory().createText();
					text.setValue(author.getInfo());
					
					r.getContent().add(text);
					newP.getContent().add(r);
					documentPart.getContent().add(index+(++i), newP);
				}
			}
			
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
			settings.setAutoHyphenation(null);
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return msg;
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
			elsevierStyle = (OOXMLElsevierStyle) publishingStyle;//new OOXMLElsevierStyle();
			if(AppEnvironment.getInstance().getNewStyleFile()!=null){
				elsevierStyle.loadStyleFromFile(AppEnvironment.getInstance().getNewStyleFile());
			} else {
				elsevierStyle.setStyleMap(AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ELSEVIER).getStyleMap());	
			}
			
			OOXMLConverterElsevier converter = new OOXMLConverterElsevier();
			OOXMLIdentifier identifier = new OOXMLIdentifier();
			identifier.init(stylePart);
			
			converter.removeDefaultHeader(documentPart);
			
			int docContentIndex = 0;
			boolean isHeadAdapted = false;
			while(docContentIndex<docContentList.size()){
				Object o = docContentList.get(docContentIndex);
				if(o instanceof P) {
					P p = (P) o;
					int headingLvl = 0;
					if(p.toString().indexOf("Shafiq ur Rehman")>=0){
						System.out.println("");
					}
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
					}
					if(identifier.isTitle(p, stylePart)){
						converter.adaptTitle(documentPart, p);
						converter.appendLayoutPara(documentPart, Context.getWmlObjectFactory().createP(), docContentIndex, 1);
						isHeadAdapted = true;
//					} else if(identifier.isSubtitle(p, stylePart)) {
//						converter.adaptSubtitle(documentPart, p);
//						converter.appendLayoutPara(documentPart, Context.getWmlObjectFactory().createP(), docContentIndex, 1);
					} else if(identifier.isAbstractHeader(p, stylePart)){
						converter.adaptAbstract(documentPart, p, docContentIndex);
						isHeadAdapted = true;
					} else if(identifier.isKeywordHeader(p, stylePart)){
						converter.adaptKeyWord(documentPart, p, docContentIndex);
						isHeadAdapted = true;
					} else if(identifier.isAcknowledgmentHeader(p, numberingPart, stylePart)){
						converter.adaptAcknowledgment(documentPart, p, docContentIndex);
						docContentIndex++;
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
							OOXMLConvertTool.replaceSectPr(p, elsevierStyle.getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT));
							OOXMLConvertTool.setAlignmentCenter(p);// fix p alignment to center
							OOXMLConvertTool.removeInd(p);
						}
						
						if(identifier.isReferencesHeaderIdentified() && !identifier.isAppendixHeaderIdentified()){
							adaptLiterature(documentPart, p);
						} else {
							OOXMLConvertTool.setPStyle(p, elsevierStyle.getStyleMap().get(AcademicFormatStructureDefinition.NORMAL).getStyleId());
						}
					}
				} else if(o instanceof SdtBlock){	// generated bibliography
					docContentIndex += OOXMLConvertTool.adaptGeneratedReferences((SdtBlock)o, elsevierStyle.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHYHEADER).getStyleId(), elsevierStyle.getStyleMap().get(AcademicFormatStructureDefinition.BIBLIOGRAPHY).getStyleId(), docContentIndex, docContentList) - 1;
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
			
			adaptFormatHeader(wordMLPackage);
			saveChange(convertedFile, wordMLPackage);
			
//			OOXMLElsevierHeaderFooterProvider.adaptHeader(AppEnvironment.getInstance().getStylePool().getStyles().get(AcademicFormatStructureDefinition.ELSEVIER).getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT), convertedFile);
			
		} catch (Exception e) {
			msg.setMessageDetails(e.getMessage());
			LogUtils.log(msg.getMessage());
		}
		
		msg.setMessageCode(UserMessage.MESSAGE_GENERAL_OK);
		LogUtils.log(msg.getMessage());
		
		return msg;
	}
	
}
