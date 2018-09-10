package lib.ooxml.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.net.URL;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.docx4j.Docx4J;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.utils.BufferUtil;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.Br;
import org.docx4j.wml.CTRel;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.Hdr;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.STLineSpacingRule;
import org.docx4j.wml.STVerticalJc;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.Style;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TblGridCol;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;

import base.AppEnvironment;
import base.UserMessage;
import db.AcademicFormatStructureDefinition;
import db.AcademicStylePool;
import tools.LogUtils;

public class OOXMLElsevierHeaderFooterProvider {

	private static ObjectFactory objectFactory = new ObjectFactory();

	private static HeaderReference headerFirstReference;
	private static HeaderReference headerDefaultReference;
	private static SectPr sectPr;
	private static AcademicStylePool stylePool;
	/*{
		sectPr = Context.getWmlObjectFactory().createSectPr();
		sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
		sectPr.getCols().setNum(BigInteger.valueOf(1));
		sectPr.getCols().setSpace(BigInteger.valueOf(720));
		sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
		sectPr.getPgMar().setTop(BigInteger.valueOf(754));
		sectPr.getPgMar().setBottom(BigInteger.valueOf(1021));
		sectPr.getPgMar().setLeft(BigInteger.valueOf(737));
		sectPr.getPgMar().setRight(BigInteger.valueOf(794));
		sectPr.getPgMar().setHeader(BigInteger.valueOf(907));
		sectPr.getPgMar().setFooter(BigInteger.valueOf(1202));
		sectPr.getPgMar().setGutter(BigInteger.valueOf(0));
		sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
		sectPr.getPgSz().setCode(BigInteger.valueOf(9));
		sectPr.getPgSz().setW(BigInteger.valueOf(10886));
		sectPr.getPgSz().setH(BigInteger.valueOf(14854));
		sectPr.setType(Context.getWmlObjectFactory().createSectPrType());
		sectPr.getType().setVal("continuous");
	}*/
	
	private static void init(){
		try {
			if(stylePool!=null){
				sectPr = stylePool.getStyles().get(AcademicFormatStructureDefinition.ELSEVIER).getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT).getPPr().getSectPr();
			}
		} catch (Exception e) {
			// TODO: handle exception
			LogUtils.log(e.getMessage());
		}
		
		if(sectPr!=null){
			
		} else {
			sectPr = Context.getWmlObjectFactory().createSectPr();
			sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			sectPr.getCols().setNum(BigInteger.valueOf(1));
			sectPr.getCols().setSpace(BigInteger.valueOf(720));
			
			sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
			sectPr.getPgMar().setTop(BigInteger.valueOf(754));
			sectPr.getPgMar().setBottom(BigInteger.valueOf(1021));
			sectPr.getPgMar().setLeft(BigInteger.valueOf(737));
			sectPr.getPgMar().setRight(BigInteger.valueOf(794));
			sectPr.getPgMar().setHeader(BigInteger.valueOf(907));
			sectPr.getPgMar().setFooter(BigInteger.valueOf(1202));
			sectPr.getPgMar().setGutter(BigInteger.valueOf(0));
			
			sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
			sectPr.getPgSz().setCode(BigInteger.valueOf(9));
			sectPr.getPgSz().setW(BigInteger.valueOf(10886));
			sectPr.getPgSz().setH(BigInteger.valueOf(14854));
		}
		
	}
	public static void main(String[] args) throws Exception {

		/*WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
		// Delete the Styles part, since it clutters up our output
		MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
		Relationship styleRel = mdp.getStyleDefinitionsPart().getSourceRelationships().get(0);
		mdp.getRelationshipsPart().removeRelationship(styleRel);
*/
		// OK, the guts of this sample:
		// The 2 things you need:
		// 1. the Header part
//		Relationship relationshipFirstPart = createHeaderPart(wordMLPackage, "header1.xml", HdrFtrRef.FIRST);
//		// 2. an entry in SectPr
//		createHeaderReference(wordMLPackage, relationshipFirstPart, HdrFtrRef.FIRST);
//		
//		Relationship relationshipDefaultPart = createHeaderPart(wordMLPackage, "header2.xml", HdrFtrRef.DEFAULT);
//		createHeaderReference(wordMLPackage, relationshipDefaultPart, HdrFtrRef.DEFAULT);
		
		
//		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File("C:/Users/Yuan Shi/Desktop/ACM Document.docx"));
		File f = new File("C:/Users/Yuan Shi/Desktop/A Framework to handle BigData for CPStestttt_ACM.docx");
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(f);
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
		try {
			
			stylePool = AppEnvironment.getInstance().getStylePool();
			if(stylePool==null){
				stylePool.loadDefinition();
			}
			init();
			
			
			/*
			P p = createTitle(wordMLPackage);
			documentPart.getContent().add(p);
			documentPart.getContent().add(createLayoutPadding(true));
			p = createTestText(wordMLPackage);
			for (int i = 0; i < 100; i++) {
				documentPart.getContent().add(p);
			}
			p = createPageLayoutPadding();
			documentPart.getContent().add(p);
			
			for(Object o : documentPart.getContent()){
				if(o instanceof P){
					P p = (P)o;
					if(p.getPPr()!=null && p.getPPr().getSectPr()!=null){
						if(true){
							if(sectPr.getPgMar()!=null){
								p.getPPr().getSectPr().setPgMar(sectPr.getPgMar());
							} else {
								LogUtils.log("ERROR");
							}
							
							if(sectPr.getPgSz()!=null){
								p.getPPr().getSectPr().setPgSz(sectPr.getPgSz());
							} else {
								LogUtils.log("ERROR2");
							}
						} else {
							p.getPPr().setSectPr(sectPr);
						}
						
					}
				}
			}
			modifyPageLayout(wordMLPackage);
			
			createHeader(wordMLPackage);
			Docx4J.save(wordMLPackage, new File("C:/Users/Yuan Shi/Desktop/OUT_HeaderFooterCreate.docx"));
			*/
			File convertedFile = new File("C:/Users/Yuan Shi/Desktop/OUT_HeaderFooterCreate.docx");
			FileUtils.copyFile(f, convertedFile);
			convertedFile.createNewFile();
			adaptHeader(stylePool.getStyles().get(AcademicFormatStructureDefinition.ELSEVIER)
					.getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT), 
					convertedFile);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		
		// Display the result as Flat OPC XML
		// FlatOpcXmlCreator worker = new FlatOpcXmlCreator(wordMLPackage);
		// worker.marshal(System.out);
		
		System.out.println("finish");
	}
	
	public static void adaptHeader(Style style, File convertedFile){
		if(style.getPPr()!=null && style.getPPr().getSectPr()!=null){
			try {
//				WordprocessingMLPackage wordMLPackage1 = WordprocessingMLPackage.load(convertedFile);
//				wordMLPackage1.save(convertedFile);
				
				WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(convertedFile);
				MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
				stylePool = new AcademicStylePool();
				stylePool.loadDefinition();
				init();
				
				
				
				createHeader(wordMLPackage);
				
				Docx4J.save(wordMLPackage, convertedFile);
			} catch (Exception e) {
				// TODO: handle exception
				LogUtils.log(e.getMessage());
			}
		} else {
			LogUtils.log("sectPr is null");
		}
	}
	
	public static void save(WordprocessingMLPackage wordMLPackage, File file){
		try {
			Docx4J.save(wordMLPackage, file);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}
	private static P createPageLayoutPadding(){
		P p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
		SectPr sectPr = p.getPPr().getSectPr();
		sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
		sectPr.getCols().setNum(BigInteger.valueOf(2));
		sectPr.getCols().setSpace(BigInteger.valueOf(720));
		sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
		sectPr.getPgSz().setW(BigInteger.valueOf(10886));
		sectPr.getPgSz().setH(BigInteger.valueOf(14854));
		sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
		sectPr.getPgMar().setTop(BigInteger.valueOf(754));
		sectPr.getPgMar().setBottom(BigInteger.valueOf(1021));
		sectPr.getPgMar().setLeft(BigInteger.valueOf(737));
		sectPr.getPgMar().setRight(BigInteger.valueOf(794));
		sectPr.getPgMar().setHeader(BigInteger.valueOf(907));
		sectPr.getPgMar().setFooter(BigInteger.valueOf(1202));
		sectPr.getPgMar().setGutter(BigInteger.valueOf(0));
		sectPr.setType(Context.getWmlObjectFactory().createSectPrType());
		sectPr.getType().setVal("continuous");
		//sectPr.setTitlePg(new BooleanDefaultTrue());
		sectPr.setFootnotePr(Context.getWmlObjectFactory().createCTFtnProps());
		sectPr.getFootnotePr().setNumFmt(Context.getWmlObjectFactory().createNumFmt());
		sectPr.getFootnotePr().getNumFmt().setVal(NumberFormat.CHICAGO);
		
		sectPr.getEGHdrFtrReferences().add(headerDefaultReference);
		
		return p;
	}
	private static P createHeaderLayoutPadding(){
		P p = Context.getWmlObjectFactory().createP();
		R r = Context.getWmlObjectFactory().createR();
		p.getContent().add(r);
		Br br = Context.getWmlObjectFactory().createBr();
		br.setType(STBrType.PAGE);
		r.getContent().add(br);
		return p;
	}
	private static P createLayoutPadding(boolean isFirst){
		P p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setSectPr(Context.getWmlObjectFactory().createSectPr());
		SectPr sectPr = p.getPPr().getSectPr();
		sectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
		sectPr.getCols().setNum(BigInteger.valueOf(1));
		sectPr.getCols().setSpace(BigInteger.valueOf(720));
		sectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
		sectPr.getPgSz().setW(BigInteger.valueOf(10886));
		sectPr.getPgSz().setH(BigInteger.valueOf(14854));
		sectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
		sectPr.getPgMar().setTop(BigInteger.valueOf(754));
		sectPr.getPgMar().setBottom(BigInteger.valueOf(1021));
		sectPr.getPgMar().setLeft(BigInteger.valueOf(737));
		sectPr.getPgMar().setRight(BigInteger.valueOf(794));
		sectPr.getPgMar().setHeader(BigInteger.valueOf(907));
		sectPr.getPgMar().setFooter(BigInteger.valueOf(1202));
		sectPr.getPgMar().setGutter(BigInteger.valueOf(0));
		sectPr.setType(Context.getWmlObjectFactory().createSectPrType());
		sectPr.getType().setVal("continuous");
		//sectPr.setTitlePg(new BooleanDefaultTrue());
		sectPr.setFootnotePr(Context.getWmlObjectFactory().createCTFtnProps());
		sectPr.getFootnotePr().setNumFmt(Context.getWmlObjectFactory().createNumFmt());
		sectPr.getFootnotePr().getNumFmt().setVal(NumberFormat.CHICAGO);
		
		sectPr.getEGHdrFtrReferences().add(headerFirstReference);
		
		return p;
	}
	private static void modifyPageLayout(WordprocessingMLPackage wordMLPackage){
		Body body = wordMLPackage.getMainDocumentPart().getJaxbElement().getBody();
		if(body!=null){
			SectPr bodySectPr = body.getSectPr();
			bodySectPr.setCols(Context.getWmlObjectFactory().createCTColumns());
			bodySectPr.getCols().setNum(BigInteger.valueOf(1));
			bodySectPr.getCols().setSpace(BigInteger.valueOf(720));
			bodySectPr.setPgSz(Context.getWmlObjectFactory().createSectPrPgSz());
			bodySectPr.getPgSz().setW(BigInteger.valueOf(10886));
			bodySectPr.getPgSz().setH(BigInteger.valueOf(14854));
			bodySectPr.setPgMar(Context.getWmlObjectFactory().createSectPrPgMar());
			bodySectPr.getPgMar().setTop(BigInteger.valueOf(754));
			bodySectPr.getPgMar().setBottom(BigInteger.valueOf(1021));
			bodySectPr.getPgMar().setLeft(BigInteger.valueOf(737));
			bodySectPr.getPgMar().setRight(BigInteger.valueOf(794));
			bodySectPr.getPgMar().setHeader(BigInteger.valueOf(907));
			bodySectPr.getPgMar().setFooter(BigInteger.valueOf(1202));
			bodySectPr.getPgMar().setGutter(BigInteger.valueOf(0));
			bodySectPr.setType(Context.getWmlObjectFactory().createSectPrType());
			bodySectPr.getType().setVal("continuous");
//			bodySectPr.setTitlePg(new BooleanDefaultTrue());
			bodySectPr.setFootnotePr(Context.getWmlObjectFactory().createCTFtnProps());
			bodySectPr.getFootnotePr().setNumFmt(Context.getWmlObjectFactory().createNumFmt());
			bodySectPr.getFootnotePr().getNumFmt().setVal(NumberFormat.CHICAGO);
		}
	}
	private static P createTestText(WordprocessingMLPackage wordMLPackage){
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
		text.setValue("Test text, it ocuppied some spaces.");
		return p;
	}
	private static P createTitle(WordprocessingMLPackage wordMLPackage){
		P p = Context.getWmlObjectFactory().createP();
		
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setJc(Context.getWmlObjectFactory().createJc());
		p.getPPr().getJc().setVal(JcEnumeration.CENTER);
		p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
		p.getPPr().getRPr().setB(new BooleanDefaultTrue());
//		p.getPPr().getRPr().setI(new BooleanDefaultTrue());
		p.getPPr().getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		p.getPPr().getRPr().getSz().setVal(BigInteger.valueOf(34));
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setB(new BooleanDefaultTrue());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(36));
		p.getContent().add(r);
		Text text = Context.getWmlObjectFactory().createText();
		r.getContent().add(text);
		text.setValue("Title");
		
		return p;
	}
	
	public static UserMessage createHeader(WordprocessingMLPackage wordMLPackage){
		UserMessage msg = new UserMessage();
		try {
			Relationship relationshipFirstPart = createHeaderPart(wordMLPackage, "header1.xml", HdrFtrRef.FIRST);
			createHeaderReference(wordMLPackage, relationshipFirstPart, HdrFtrRef.FIRST);
					
			Relationship relationshipDefaultPart = createHeaderPart(wordMLPackage, "header2.xml", HdrFtrRef.DEFAULT);
			createHeaderReference(wordMLPackage, relationshipDefaultPart, HdrFtrRef.DEFAULT);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
				
		return msg;
	}

	public static Relationship createHeaderPart(WordprocessingMLPackage wordprocessingMLPackage, String headerPartName, HdrFtrRef refType) throws Exception {

//		HeaderPart headerPart = new HeaderPart(new PartName("/word/header/"+headerPartName));
		HeaderPart headerPart = new HeaderPart(new PartName("/word/"+headerPartName));
		Relationship rel = wordprocessingMLPackage.getMainDocumentPart().addTargetPart(headerPart);

		// After addTargetPart, so image can be added properly
		if(refType!=null){
			if(refType.equals(HdrFtrRef.DEFAULT)) {
				headerPart.setJaxbElement(getDefaultHdr(wordprocessingMLPackage, headerPart));
			} else if(refType.equals(HdrFtrRef.FIRST)) {
				headerPart.setJaxbElement(getFirstHdr(wordprocessingMLPackage, headerPart));
			} else {
				headerPart.setJaxbElement(getFirstHdr(wordprocessingMLPackage, headerPart));
			}
		} else {
			headerPart.setJaxbElement(getDefaultHdr(wordprocessingMLPackage, headerPart));
		}
		

		return rel;
	}

	public static void createHeaderReference(WordprocessingMLPackage wordprocessingMLPackage, Relationship relationship, HdrFtrRef refType)
			throws InvalidFormatException {

		List<SectionWrapper> sections = wordprocessingMLPackage.getDocumentModel().getSections();

		SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
		// There is always a section wrapper, but it might not contain a sectPr
		if (sectPr == null) {
			sectPr = OOXMLElsevierHeaderFooterProvider.sectPr;//objectFactory.createSectPr();
			wordprocessingMLPackage.getMainDocumentPart().addObject(sectPr);
			sections.get(sections.size() - 1).setSectPr(sectPr);
		}
		
		if(refType.equals(HdrFtrRef.FIRST)){
//			System.out.println(
//					sections.get(0).getPageDimensions().getWritableWidthTwips()
//					+"\n"
//					+sections.get(0).getPageDimensions().getHeaderMargin()
//					+"\n"
//					+sections.get(0).getPageDimensions().getPgMar().getLeft()
//					
//					);
//			System.out.println(
//			sections.get(sections.size() - 1).getPageDimensions().getWritableWidthTwips()
//			+"\n"
//			+sections.get(sections.size() - 1).getPageDimensions().getHeaderMargin()
//			+"\n"
//			+sections.get(sections.size() - 1).getPageDimensions().getPgMar().getLeft()
//			);
			
		}
		
		HeaderReference headerReference = objectFactory.createHeaderReference();
		headerReference.setId(relationship.getId());
//		headerReference.setType(HdrFtrRef.DEFAULT);
		headerReference.setType(refType);
		sectPr.getEGHdrFtrReferences().add(headerReference);// add header or
		// footer references
		if(refType.equals(HdrFtrRef.FIRST)){
			sectPr.setTitlePg(new BooleanDefaultTrue());
			headerFirstReference = headerReference;
		} else if(refType.equals(HdrFtrRef.DEFAULT)) {
			headerDefaultReference = headerReference;
		} else {
			// EVEN
		}
	}

	public static Hdr getFirstHdr(WordprocessingMLPackage wordprocessingMLPackage, Part sourcePart) throws Exception {

		Hdr hdr = objectFactory.createHdr();

//		File file = new File(System.getProperty("user.dir") + "/src/sources/elsevier/elsevier_header_image1.png");
//		java.io.InputStream is = new java.io.FileInputStream(file);

//		hdr.getContent().add(newImage(wordprocessingMLPackage, sourcePart, BufferUtil.getBytesFromInputStream(is),
//				"filename", "alttext", 1, 2));
		hdr.getContent().add(createDocHeaderTbl(wordprocessingMLPackage, sourcePart));
		hdr.getContent().add(createHeaderLayoutPadding());
		return hdr;

	}
	
	public static Hdr getDefaultHdr(WordprocessingMLPackage wordprocessingMLPackage, Part sourcePart) throws Exception {

		Hdr hdr = objectFactory.createHdr();

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

		hdr.getContent().add(p);
		hdr.getContent().add(createHeaderLayoutPadding());
		return hdr;

	}

	// public static P getP() {
	// P headerP = objectFactory.createP();
	// R run1 = objectFactory.createR();
	// Text text = objectFactory.createText();
	// text.setValue("123head123");
	// run1.getRunContent().add(text);
	// headerP.getParagraphContent().add(run1);
	// return headerP;
	// }

	public static org.docx4j.wml.P newImage(WordprocessingMLPackage wordMLPackage, Part sourcePart, byte[] bytes,
			String filenameHint, String altText, int id1, int id2, long cx, long cy) throws Exception {

		BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, sourcePart, bytes);

		Inline inline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
		if(cx>0 && cy>0){
			inline.getExtent().setCx(cx);
			inline.getExtent().setCy(cy);
		}
		// Now add the inline in w:p/w:r/w:drawing
		org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory();
		org.docx4j.wml.P p = factory.createP();
		org.docx4j.wml.R run = factory.createR();
		p.getContent().add(run);
		org.docx4j.wml.Drawing drawing = factory.createDrawing();
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);

		return p;

	}
	
	private static Tbl createDocHeaderTbl(WordprocessingMLPackage wordMLPackage, Part headerPart){
		Tbl tbl = null;
		
		try {
			int writableWidthTwips = 9360;	//wordMLPackage.getDocumentModel().getSections().get(0).getPageDimensions().getWritableWidthTwips();
			tbl = TblFactory.createTable(1, 3, writableWidthTwips);// in 1 row, 3 columns, auto width
			
			tbl.setTblGrid(Context.getWmlObjectFactory().createTblGrid());
			
			int[] colWidth = {1688, 5760, 1912};//{1200, writableWidthTwips-1200-1920, 1920};
			
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
								
//								imageURL = getClass().getResource("/sources/elsevier/elsevier_header_image1.png");
//								imageFile = new File(imageURL.toURI());
//								imageFile = new File(System.getProperty("user.dir") + "/src/sources/elsevier/elsevier_header_image1.png");
								
								InputStream in = System.getProperty("user.dir").getClass().getResourceAsStream("/sources/elsevier/elsevier_header_image1.png");
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
//								imageFile = new File("src/sources/elsevier/elsevier_header_image1.png");
								bytes = convertImageToByteArray(imageFile);
								imageFile.delete();
//								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Elsevier", "Elsevier logo", 604800, 669600);
								
								p = Context.getWmlObjectFactory().createP();
								R rr = Context.getWmlObjectFactory().createR();
//								rr.getContent().add(newImage(wordMLPackage, headerPart, bytes, "Elsevier", "Elsevier logo", 1, 2, 604800, 669600));
								rr.getContent().add(newImage(wordMLPackage, headerPart, bytes, "Elsevier", "Elsevier logo", 1, 1, 685800, 754380));
								p.getContent().add(rr);
								
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
								
//								imageURL = getClass().getResource("/sources/elsevier/elsevier_header_cs_image3.png");
//								imageFile = new File(imageURL.toURI());
//								imageFile = new File(System.getProperty("user.dir") + "/src/sources/elsevier/elsevier_header_cs_image3.png");
								InputStream in2 = System.getProperty("user.dir").getClass().getResourceAsStream("/sources/elsevier/elsevier_header_cs_image3.png");
								imageFile = File.createTempFile("img1", ".png");
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
								imageFile.delete();
//								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Elsevier", "Elsevier logo", 604800, 669600);
								
								p = Context.getWmlObjectFactory().createP();
								R r3 = Context.getWmlObjectFactory().createR();
//								r3.getContent().add(newImage(wordMLPackage, headerPart, bytes, "Procedia", "Procedia logo", 1, 2, 1137600, 741600));
								r3.getContent().add(newImage(wordMLPackage, headerPart, bytes, "Procedia", "Procedia logo", 2, 2, 2156460, 472440));
								p.getContent().add(r3);
//								bytes = convertImageToByteArray(imageFile);
//								p = addImageToPackage(wordMLPackage, bytes, headerPart, "Procedia", "Procedia logo", 1137600, 741600);
								
								tc.getContent().add(p);
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
	
	private static P addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes, HeaderPart part, String fileNameHint, String altText) {
		P paragraph = null;
		
		addImageToPackage(wordMLPackage, bytes, part, fileNameHint, altText, 0, 0);
				
		return paragraph;
	}
	private static P addImageToPackage(WordprocessingMLPackage wordMLPackage, byte[] bytes, HeaderPart part, String fileNameHint, String altText, long cx, long cy) {
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

			wordMLPackage.getMainDocumentPart().addObject(paragraph);

		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
			LogUtils.log(e.getMessage());
		}
		return paragraph;
	}

	private static P addInlineImageToParagraph(Inline inline) {
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
}
