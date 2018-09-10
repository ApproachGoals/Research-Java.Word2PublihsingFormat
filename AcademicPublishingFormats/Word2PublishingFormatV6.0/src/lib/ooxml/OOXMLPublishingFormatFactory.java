package lib.ooxml;

import java.io.File;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;

import base.AppEnvironment;
import base.UserMessage;
import lib.AbstractPublishingFormatFactory;
import lib.AbstractPublishingStyle;
import lib.ooxml.process.OOXMLACMConvertingFactory;
import lib.ooxml.process.OOXMLConvertingInterface;
import lib.ooxml.process.OOXMLElsevier1ConvertingFactory;
import lib.ooxml.process.OOXMLElsevier2ConvertingProvider;
import lib.ooxml.process.OOXMLIEEEConvertingFactory;
import lib.ooxml.process.OOXMLSpringerConvertingFactory;
import lib.ooxml.style.OOXMLACMStyle;
import lib.ooxml.style.OOXMLElsevierStyle;
import lib.ooxml.style.OOXMLElsevierStyle2;
import lib.ooxml.style.OOXMLIEEEStyle;
import lib.ooxml.style.OOXMLSpringerStyle;
import tools.LogUtils;

public class OOXMLPublishingFormatFactory extends AbstractPublishingFormatFactory{

	private AbstractPublishingStyle publishingStyle;
	
	public AbstractPublishingStyle getPublishingStyle() {
		return publishingStyle;
	}

	public void setPublishingStyle(AbstractPublishingStyle publishingStyle) {
		this.publishingStyle = publishingStyle;
	}

	public UserMessage convertProcess(File convertedFile) {
		UserMessage msg = new UserMessage();
		if(convertedFile!=null){
			try {
				if(convertedFile.canWrite()){
					
				} else {
					LogUtils.log("The goal file cannot be written, please make sure the file lock is released.");
				}
				WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(convertedFile);
				LogUtils.log("original file was loaded into wordMLPackage.");
				MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
				LogUtils.log("finish to get MainDocumentPart from wordMLPackage.");
				StyleDefinitionsPart stylePart = documentPart.getStyleDefinitionsPart();
				LogUtils.log("get the StyleDefinitionsPart from documentPart");
//				DocumentSettingsPart settingPart = documentPart.getDocumentSettingsPart();
//				LogUtils.log("get the documentSettingPart from documentPart");
				
//				List<Style> newPublishingStyleList = null;
				OOXMLConvertingInterface converterProvider = null;
				if(publishingStyle instanceof OOXMLACMStyle) {
//					newPublishingStyleList = ((OOXMLACMStyle)publishingStyle).getStyleList();
//					converterProvider = new OOXMLACMConvertingProvider(convertedFile, wordMLPackage, documentPart, stylePart);
//					converterProvider.convert2Format(convertedFile, wordMLPackage, documentPart, stylePart);
					new Thread(new OOXMLACMConvertingFactory(convertedFile, wordMLPackage, documentPart, stylePart, publishingStyle)).start();
				} else if(publishingStyle instanceof OOXMLIEEEStyle) {
//					newPublishingStyleList = ((OOXMLIEEEStyle)publishingStyle).getStyleList();
//					converterProvider = new OOXMLIEEEConvertingProvider();
//					converterProvider.convert2Format(convertedFile, wordMLPackage, documentPart, stylePart);
					new Thread(new OOXMLIEEEConvertingFactory(convertedFile, wordMLPackage, documentPart, stylePart, publishingStyle)).start();
				} else if(publishingStyle instanceof OOXMLSpringerStyle) {
//					newPublishingStyleList = ((OOXMLSpringerStyle)publishingStyle).getStyleList();
//					converterProvider = new OOXMLSpringerConvertingProvider();
//					converterProvider.convert2Format(convertedFile, wordMLPackage, documentPart, stylePart);
					new Thread(new OOXMLSpringerConvertingFactory(convertedFile, wordMLPackage, documentPart, stylePart, publishingStyle)).start();
				} else if(publishingStyle instanceof OOXMLElsevierStyle) {
//					newPublishingStyleList = ((OOXMLElsevierStyle)publishingStyle).getStyleList();
//					converterProvider = new OOXMLElsevier1ConvertingProvider();
//					converterProvider.convert2Format(convertedFile, wordMLPackage, documentPart, stylePart);
					new Thread(new OOXMLElsevier1ConvertingFactory(convertedFile, wordMLPackage, documentPart, stylePart, publishingStyle)).start();
				} else if(publishingStyle instanceof OOXMLElsevierStyle2){
//					newPublishingStyleList = ((OOXMLElsevierStyle2)publishingStyle).getStyleList();
					converterProvider = new OOXMLElsevier2ConvertingProvider();
					converterProvider.convert2Format(convertedFile, wordMLPackage, documentPart, stylePart);
				} else {
					msg.setMessageDetails("This format was not implemented.");
				}
				
			} catch (Exception e) {
				// TODO: handle exception
				msg.setMessageDetails(e.getMessage());
			}
		}
		
		return msg;
	}
	
}
