package lib.odf;

import java.io.File;

import base.UserMessage;
import lib.AbstractPublishingFormatFactory;
import lib.AbstractPublishingStyle;

public class ODFPublishingFormatFactory extends AbstractPublishingFormatFactory {

	private AbstractPublishingStyle publishingStyle;
	
	public AbstractPublishingStyle getPublishingStyle() {
		return publishingStyle;
	}

	public void setPublishingStyle(AbstractPublishingStyle publishingStyle) {
		this.publishingStyle = publishingStyle;
	}

	@Override
	public UserMessage convertProcess(File convertedFile) {
		// TODO Auto-generated method stub
		return null;
	}

}
