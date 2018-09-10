package lib.ooxml.process;

import java.io.File;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.SectPr;

import base.UserMessage;
import lib.ooxml.author.Authors;

public interface OOXMLConvertingInterface {

	public UserMessage adaptAbstract(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptAuthors(MainDocumentPart documentPart, int docContentIndex, int authorCount, Authors authors);
	
	public UserMessage adaptCaption(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptFontStyleClass(File convertedFile, WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart, StyleDefinitionsPart stylePart);
	
	public UserMessage adaptFootNote(File convertedFile, WordprocessingMLPackage wordMLPackage);
	
	public UserMessage adaptHeading1(P p);
	
	public UserMessage adaptHeading2(P p);
	
	public UserMessage adaptHeading3(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptHeading4(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptImages(MainDocumentPart documentPart);
	
	public UserMessage adaptKeyWord(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptLiterature(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptLiteratureHeader(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptNumberFile(MainDocumentPart documentPart);
	
	public UserMessage adaptPageLayout(MainDocumentPart documentPart, WordprocessingMLPackage wordMLPackage);
	
	public UserMessage adaptSettings(MainDocumentPart documentPart);
	
	public UserMessage adaptTitle(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptTable(MainDocumentPart documentPart);
	
	public UserMessage adaptTableCaption(MainDocumentPart documentPart, P p);

	public UserMessage appendLayoutPara(MainDocumentPart documentPart, P p, int docContentIndex, int cols);
	
	public UserMessage convert2Format(File convertedFile, WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart, StyleDefinitionsPart stylePart);
	
	public UserMessage removePSetting(PPr pPr);
	
	public UserMessage removeUselessPageBreaks(WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart);
	
	public UserMessage saveChange(File convertedFile, WordprocessingMLPackage wordMLPackage);
	
}
