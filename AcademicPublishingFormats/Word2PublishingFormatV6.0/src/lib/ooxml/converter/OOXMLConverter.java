package lib.ooxml.converter;

import java.io.File;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;

import base.UserMessage;
import lib.ooxml.author.Authors;
import lib.ooxml.identifier.OOXMLIdentifier;

public interface OOXMLConverter {
	
	public UserMessage adaptAbstract(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptAcknowledgment(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptAppendixHeader(MainDocumentPart documentPart, P p, int docContentIndex);
	
	public UserMessage adaptAuthors(MainDocumentPart documentPart, int docContentIndex, int authorCount, Authors authors);
	
	public UserMessage adaptFigureCaption(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptFontStyleClass(File convertedFile, WordprocessingMLPackage wordMLPackage, MainDocumentPart documentPart, StyleDefinitionsPart stylePart);
	
	public UserMessage adaptFootNote(File convertedFile, WordprocessingMLPackage wordMLPackage);
	
	public UserMessage adaptGeneratedReferences();
	
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

	public UserMessage adaptSubtitle(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptTitle(MainDocumentPart documentPart, P p);
	
	public UserMessage adaptTable(MainDocumentPart documentPart);
	
	public UserMessage adaptTableCaption(MainDocumentPart documentPart, P p);

	public UserMessage appendLayoutPara(MainDocumentPart documentPart, P p, int docContentIndex, int cols);
	
	public UserMessage removeDefaultHeader(MainDocumentPart documentPart);
}
