package lib;

import java.io.File;

import base.UserMessage;

public abstract class AbstractPublishingFormatFactory {

	public abstract UserMessage convertProcess(File convertedFile);

}
