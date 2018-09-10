package lib;

import lib.odf.ODFPublishingFormatFactory;
import lib.odf.style.ODFACMStyle;
import lib.ooxml.OOXMLPublishingFormatFactory;
import lib.ooxml.style.OOXMLACMStyle;
import lib.ooxml.style.OOXMLElsevierStyle;
import lib.ooxml.style.OOXMLElsevierStyle2;
import lib.ooxml.style.OOXMLIEEEStyle;
import lib.ooxml.style.OOXMLSpringerStyle;

public class FormatConverter {

	public static String FILE_FORMAT_MODE_OOXML_ACM = "file.format.mode.ooxml.acm";
	public static String FILE_FORMAT_MODE_OOXML_ELSEVIER2 = "file.format.mode.ooxml.elsevier2";
	public static String FILE_FORMAT_MODE_OOXML_IEEE = "file.format.mode.ooxml.ieee";
	public static String FILE_FORMAT_MODE_OOXML_SPRINGER = "file.format.mode.ooxml.springer";
	public static String FILE_FORMAT_MODE_OOXML_ELSEVIER1 = "file.format.mode.ooxml.elsevier1";
	public static String FILE_FORMAT_MODE_ODF_ACM = "file.format.mode.odf.acm";
	
	public static AbstractPublishingFormatFactory getFactory(String mode){
		AbstractPublishingFormatFactory factory = null;
		
		if(FILE_FORMAT_MODE_OOXML_ACM.equals(mode)){
			factory = new OOXMLPublishingFormatFactory();
			((OOXMLPublishingFormatFactory) factory).setPublishingStyle(new OOXMLACMStyle());
		} else if(FILE_FORMAT_MODE_OOXML_ELSEVIER1.equals(mode)){
			factory = new OOXMLPublishingFormatFactory();
			((OOXMLPublishingFormatFactory) factory).setPublishingStyle(new OOXMLElsevierStyle());
		} else if(FILE_FORMAT_MODE_OOXML_ELSEVIER2.equals(mode)){
			factory = new OOXMLPublishingFormatFactory();
			((OOXMLPublishingFormatFactory) factory).setPublishingStyle(new OOXMLElsevierStyle2());
		} else if(FILE_FORMAT_MODE_OOXML_IEEE.equals(mode)) {
			factory = new OOXMLPublishingFormatFactory();
			((OOXMLPublishingFormatFactory) factory).setPublishingStyle(new OOXMLIEEEStyle());
		} else if(FILE_FORMAT_MODE_OOXML_SPRINGER.equals(mode)){
			factory = new OOXMLPublishingFormatFactory();
			((OOXMLPublishingFormatFactory) factory).setPublishingStyle(new OOXMLSpringerStyle());
		} else if(FILE_FORMAT_MODE_ODF_ACM.equals(mode)) {
			factory = new ODFPublishingFormatFactory();
			((ODFPublishingFormatFactory) factory).setPublishingStyle(new ODFACMStyle());
		}
		
		return factory;
	}
}
