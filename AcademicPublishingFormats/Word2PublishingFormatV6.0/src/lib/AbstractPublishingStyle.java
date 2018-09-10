package lib;

import java.io.File;
import java.util.Map;

import org.docx4j.wml.Style;

public interface AbstractPublishingStyle {
	
	public static String STYLENAME_ABSTRACT = "stylename.abstract";
	public static String STYLENAME_ABSTRACTHEADER = "stylename.abstractheader";
	public static String STYLENAME_AFFILIATIONS = "stylename.affiliations";
	public static String STYLENAME_AUTHORS = "stylename.authors";
	public static String STYLENAME_CAPTION = "stylename.caption";
	public static String STYLENAME_DEFAULTPARAGRAPH = "stylename.defaultparagraph";
	public static String STYLENAME_EMAIL = "stylename.email";
	public static String STYLENAME_FIGURE = "stylename.figure";
	public static String STYLENAME_FIGURECHAR = "stylename.figureChar";
	public static String STYLENAME_FOOTNOTES = "stylename.footnotes";
	public static String STYLENAME_HEADING1 = "stylename.heading1";
	public static String STYLENAME_HEADING2 = "stylename.heading2";
	public static String STYLENAME_HEADING3 = "stylename.heading3";
	public static String STYLENAME_HEADING3CHAR = "stylename.heading3Char";
	public static String STYLENAME_HEADING4 = "stylename.heading4";
	public static String STYLENAME_HEADING4CHAR = "stylename.heading4Char";
	public static String STYLENAME_HEADING5 = "stylename.haeding5";	// from IEEE
	public static String STYLENAME_HYPERLINK = "stylename.hyperlink";
	public static String STYLENAME_KEYWORD = "stylename.keyword";
	public static String STYLENAME_KEYWORDHEAD = "stylename.keywordhead";
	public static String STYLENAME_PHONE = "stylename.phone";
	public static String STYLENAME_REFERENCE = "stylename.reference";
	public static String STYLENAME_REFERENCEHEAD = "stylename.referencehead";
	public static String STYLENAME_SUBTITLE = "stylename.subtitle";// acm not used, from IEEE
	public static String STYLENAME_TABLEHEADER = "stylename.tableheader";// acm not used, from IEEE
	public static String STYLENAME_TABLEHEADERCHAR = "stylename.tableheaderChar";
	public static String STYLENAME_TABLECOLHEADER = "stylename.tablecolheadeer";// from IEEE
	public static String STYLENAME_TITLE = "stylename.title";
	
	public void loadStyleFromFile(File file);
}
