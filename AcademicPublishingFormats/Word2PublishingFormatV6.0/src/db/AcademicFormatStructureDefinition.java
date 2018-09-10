package db;

public class AcademicFormatStructureDefinition {

	/*
	 * names of common academic document's components
	 */
	public static String TITLE = "Title";
	public static String SUBTITLE = "SubTitle";
	public static String AUTHOR = "Authors";
	public static String AFFILIATION = "Affiliation";
	public static String DATE = "Date";
	public static String EMAIL = "Email";
	public static String ABSTRACT = "Abstract";
	public static String ABSTRACTHEADER = "AbstractHeader";
	public static String EPIGRAPH = "Epigraph";
	public static String EPICITE = "EpiCite";
	public static String BLOCKQUOTE = "BlockQuote";
	public static String BLOCKQUOTECITE = "BlockQuoteCite";
	public static String KEYWORDS = "keyword";
	public static String KEYWORDHEADER = "KeywordHeader";
	public static String TERMS = "Terms";
	public static String HEADING1 = "Heading1";
	public static String HEADING2 = "Heading2";
	public static String HEADING3 = "Heading3";
	public static String HEADING3CHAR = "Heading3Char";
	public static String HEADING4 = "Heading4";
	public static String HEADING4CHAR = "Heading4Char";
	public static String HEADING5 = "Heading5";
	public static String HEADING5CHAR = "Heading5Char";
	public static String LISTBULLET = "ListBullet";
	public static String LISTNUMBER = "ListNumber";
	public static String LISTLABEL = "ListLabel";
	public static String LISTCONTINUE = "ListContinue";
	public static String NORMAL = "Normal";
	public static String NORMALCHAR = "NormalChar";
	public static String FIGURE = "Figure";
	public static String FIGUREHEADER = "FigureHeader";
	public static String TABLE = "Table";
	public static String CAPTION = "Caption";
	public static String CAPTIONHEADER = "CaptionHeader";
	public static String ACKNOWLEDGMENT = "Acknowledgment";
	public static String ACKNOWLEDGMENTHEADER = "AcknowledgmentHeader";
	public static String ACKNOWLEDGMENTCHAR = "AcknowledgmentChar";
	public static String APPENDIX = "Appendix";
	public static String APPENDIXHEADER = "AppendixHeader";
	public static String BIBLIOGRAPHY = "Bibliography";
	public static String BIBLIOGRAPHYHEADER = "BibliographyHeader";
	public static String CITATION = "Citation";
	public static String FOOTNOTE = "Footnotes";
	public static String EMPHASIS = "Emphasis";
	public static String PAGEFORMAT = "Pageformat";
	public static String FIRSTPARA = "p1a";
	public static String firstname = "Firstname";
	public static String surname = "Surname";
	public static String RPRDEFAULT = "rPrDefault";
	/*
	 * definition of JSON key for OOXML style
	 */
	public static String DEFAULTSYLE = "defaultsyle";	// String ["true", "false"], if value is equals true, this style is default style
	public static String FONTFAMILY = "fontname";		// String, for the font name
	public static String FONTSIZE = "fontsz";			// String, which will be converted into BigInteger, for font size
	public static String FONTBOLD = "bold";				// String ["true", "false"], if value is equals true, font weight is bold.
	public static String FONTITATLIC = "italic";		// String ["true", "false"], if value is equals true, font italic is bold.
	public static String FONTCOLOR = "fontcolor";		// String "six hexadecimal", for font colors.
	public static String FONTCAPS = "caps";				// String ["true", "false"], if value is equals true, font caps (upperCase).
	public static String FONTSMALLCAPS = "smallcaps";	// String ["true", "false"], if value is equals true, font smallCaps (upperCase, and first character has bigger size).
	public static String FONTUNDERLINE = "underline";	// String ["true", "false"], if value is equals true, text with underline.
	public static String PARABEFORE = "spacingbefore";	// String "integer", space before paragraph.
	public static String PARAAFTER = "spacingafter";	// String "integer", space after paragraph.
	public static String PARALINE = "spacingline";		// String "integer", line space of paragraph.
	public static String PARALINERULE = "linerule";		// String ["auto", "exact", "atLeast"], mode for line space
	public static String PARAKEEPLINE = "keepline";		// String ["true", "false"], if value is equals true, text keep in line possibly.
	public static String PARAKEEPNEXT = "keepnext";		// String ["true", "false"], if value is equals true, text and next paragraph keep possibly in one page.
	public static String PARANEXT = "next";				// String, the default style of next paragraph.
	public static String PARAJC = "jc";					// String ["center", "left", "right", "both"], the alignment of the paragraph
	public static String PARALEFT = "indleft";			// String "integer", left indentation of paragraph
	public static String PARARIGHT = "indight";			// String "integer", right indentation of paragraph
	public static String PARAFIRSTLINE = "indfirstline";// String "integer", indentation of first line in paragraph
	public static String PARAHANGING = "indhanging";	// String "integer", indentation of line without first line in paragraph
	public static String BASEDON = "basedon";			// String "styleId", inherited from the style, which its style id is referred. 
	public static String LINK = "link";					// String "styleId", link to other style, for example, combination of paragraph and character. 
	public static String TYPE = "type";					// String ["paragraph", "character"]
	public static String ID = "id";						// String, style id of the style
	public static String NAME = "name";					// String, style name
	public static String PARASECTPGSZW = "pagesizewidth";				// String "integer", page width
	public static String PARASECTPGSZH = "pagesizeheight";				// String "integer", page height
	public static String PARASECTPGSZCODE = "pagesizecode";
	public static String PARASECTPGMARGINTOP = "pagemargintop";			// String "integer", distance between top edge of page and text top edge
	public static String PARASECTPGMARGINBOTTOM = "pagemarginbottom";	// String "integer", distance between bottom edge of page and text bottom edge
	public static String PARASECTPGMARGINLEFT = "pagemarginleft";		// String "integer", distance between the left edge of the page and the left edge of the text
	public static String PARASECTPGMARGINRIGHT = "pagemarginright";		// String "integer", distance between the right edge of the page and the right edge of the text
	public static String PARASECTPGGUTTER = "pagegutter";			// String "integer", space size of gutter
	public static String PARASECTPGHEADER = "pageheader";			// String "integer", distance between top edge of page and header top edge
	public static String PARASECTPGFOOTER = "pagefooter";			// String "integer", distance between bottom edge of page and footer bottom edge
	public static String PARASECTCOLNUM = "pagecolumnnumber";		// String "integer", column number of page
	public static String PARASECTCOLSPACING = "pagecolumnspacing";	// String "integer", space among columns
	public static String PARASECTLINEPITCH = "linePitch";			
	public static String PARANUMLIST = "numid";						// String, numbering id
	public static String PARALVL = "lvl";							// String, lvl value
	public static String PARABDRTOPTYPE = "parageraphbordertoptype";// String, paragraph top border, ["single", "dot_dash", "double"]
	public static String PARABDRTOPSZ = "parageraphbordertopsize";	// String, top border line size
	public static String PARABDRTOPSPACE = "parageraphbordertopspace";	// String, border top space
	public static String PARABDRTOPCOLOR = "parageraphbordertopcolor";	// String, border top color
	public static String PARABDRBOTTOMTYPE = "parageraphborderbottomtype";	// similar to border top
	public static String PARABDRBOTTOMSZ = "parageraphborderbottomsize";	// similar to border top
	public static String PARABDRBOTTOMSPACE = "parageraphborderbottomspace";// similar to border top
	public static String PARABDRBOTTOMCOLOR = "parageraphborderbottomcolor";// similar to border top
	public static String PARATABVAL = "tabval";						// tab spacing mode (e.g. left)
	public static String PARATABPOS = "tabpos";						// tab spacing position (number)
	public static String SHD_VAL = "shdval";						// shadow value
	public static String SHD_COLOR = "shdcolor";					// shadow fill type
	public static String SHD_FILL = "shdfill";						// shadow fill color
	public static String NUMBERINGFMT = "numfmt";							// String ["decimal", "upperRoman", "lowerLetter", "upperLetter"], numbering format
	public static String NUMBERINGLVLTEXT = "lvltext";						// String "%1.", "%2"|"%1.%2", string pattern
	public static String NUMBERING_PPR_IND_LEFT = "num_ppr_ind_left";		// String, numbering indentation, similar to paragraph indentation
	public static String NUMBERING_PPR_IND_HANGING = "num_ppr_ind_hanging";	// String, numbering indentation, similar to paragraph indentation
	public static String NUMBERING_PPR_IND_FIRSTLINE = "num_ppr_ind_firstline";// String, numbering indentation, similar to paragraph indentation
	
	public static String NUMBERINGSECTIONSTRING = "sectionNumber";			// String, help definition, numbering mode for section
	public static String NUMBERINGREFSTRING = "refNumber";					// String, help definition, numbering mode for reference
	public static String NUMBERINGCAPTIONSTRING = "captionNumber";			// String, help definition, numbering mode for caption
	public static String NUMBERINGFIGURESTRING = "figureNumber";			// String, help definition, numbering mode for figure
	
	public static String ACM = "acm";
	public static String IEEE = "ieee";
	public static String SPRINGER = "springer";
	public static String ELSEVIER = "elsevier";
}
