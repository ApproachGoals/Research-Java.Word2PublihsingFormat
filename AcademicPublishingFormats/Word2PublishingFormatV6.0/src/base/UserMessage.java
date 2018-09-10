package base;

public class UserMessage {

	public static String MESSAGE_TITLE_ADAPT_FINISH = "title.adapt.finish";
	public static String MESSAGE_ABSTRACT_ADAPT_FINISH = "abstract.adapt.finish";
	public static String MESSAGE_ACKNOWLEDGMENT_ADAPT_FINISH = "acknowledgment.adapt.finish";
	public static String MESSAGE_AUTHORS_ADAPT_FINISH = "authors.adapt.finish";
	public static String MESSAGE_BIBLIOGRAPHY_HEADER_ADAPT_FINISH = "bibliography.header.adapt.finish";
	public static String MESSAGE_CAPTION_ADAPT_FINISH = "capton.adapt.finish";
	public static String MESSAGE_FOOTNOTE_ADAPT_FINISH = "footnote.adapt.finish";
	public static String MESSAGE_KEYWORD_ADAPT_FINISH = "keyword.adapt.finish";
	public static String MESSAGE_HEADING1_ADAPT_FINISH = "heading1.adapt.finish";
	public static String MESSAGE_HEADING2_ADAPT_FINISH = "heading2.adapt.finish";
	public static String MESSAGE_HEADING3_ADAPT_FINISH = "heading3.adapt.finish";
	public static String MESSAGE_HEADING4_ADAPT_FINISH = "heading4.adapt.finish";
	public static String MESSAGE_IMAGE_ADAPT_FINISH = "images.adapt.finish";
	public static String MESSAGE_NEWLINE_LAYOUT_APPENDED = "pagelayout.newline.appended";
	public static String MESSAGE_NUMBERFILE_ADAPT_FINISH = "numberfile.adapt.finish";
	public static String MESSAGE_PAGELAYOUT_ADAPT_FINISH = "pagelayout.adapt.finish";
	public static String MESSAGE_REFERENCEHEADER_ADPAT_FINISH = "referenceheader.adapt.finish";
	public static String MESSAGE_REFERENCE_ADPAT_FINISH = "reference.adapt.finish";
	public static String MESSAGE_TABLECAPTION_ADAPT_FINISH = "tablecaption.adapt.finish";
	public static String MESSAGE_STYLECLASS_ADAPT_FINISH = "styleclass.adapt.finish";
	
	public static String MESSAGE_GENERAL_OK = "0";
	
	public static String MESSAGE_GENERAL_ERROR = "-1";
	
	public static String MESSAGE_ERROR_NO_CLASS_ID = "error.no.classid";
	
	public static String MESSAGE_UNABLESAVE_ERROR = "unable_saved";
	
	public static String MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM = "new.style.create.error";
	
	public static String MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM_NULL = "new.style.item.null";
	
	public static String MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM_NONE_CID = "new.style.item.no.clzid";
	
	public static String MESSAGE_UNABLE_FOUND_STYLE_BY_CLZID = "style.notfound.byid";
	
	public static String CHKRESULT_NULL = "result.null";
	public static String CHKRESULT_UNACCEPTED = "result.unaccepted";
	public static String CHKRESULT_ACCEPTED = "result.accepted";
	
	private int beginPosition;
	private int endPosition;
	private int indexOfLastR;
	private String sameStyleName;
	private boolean isFullySuitable;
	
	private boolean chkresultSuitable;
	
	private String messageCode;
	
	private String messageDetails;
	
	public boolean isChkresultSuitable() {
		return chkresultSuitable;
	}

	public void setChkresultSuitable(boolean chkresultSuitable) {
		this.chkresultSuitable = chkresultSuitable;
	}

	public String getMessageCode() {
		return messageCode;
	}

	public void setMessageCode(String messageCode) {
		this.messageCode = messageCode;
	}

	public String getMessageDetails() {
		return messageDetails;
	}

	public void setMessageDetails(String messageDetails) {
		this.messageDetails = messageDetails;
	}

	public int getIndexOfLastR() {
		return indexOfLastR;
	}

	public void setIndexOfLastR(int indexOfLastR) {
		this.indexOfLastR = indexOfLastR;
	}

	public int getBeginPosition() {
		return beginPosition;
	}

	public void setBeginPosition(int beginPosition) {
		this.beginPosition = beginPosition;
	}

	public boolean isFullySuitable() {
		return isFullySuitable;
	}

	public void setFullySuitable(boolean isFullySuitable) {
		this.isFullySuitable = isFullySuitable;
	}

	public int getEndPosition() {
		return endPosition;
	}

	public void setEndPosition(int endPosition) {
		this.endPosition = endPosition;
	}

	public String getSameStyleName() {
		return sameStyleName;
	}

	public void setSameStyleName(String sameStyleName) {
		this.sameStyleName = sameStyleName;
	}

	public UserMessage(){
		
	}
	public UserMessage(String errorCode){
		setMessageCode(errorCode);
	}
	public UserMessage(String errorCode, String messageDetails){
		setMessageCode(errorCode);
		setMessageDetails(messageDetails);
	}
	
	public String getMessage() {
		String result = "";
		if(getMessageCode()!=null){
			if(getMessageCode().equals(MESSAGE_ERROR_NO_CLASS_ID)) {
				result = "error, the style's classid was not defined";
			} else if(getMessageCode().equals(MESSAGE_GENERAL_ERROR)){
				result = "some logic error happend.";
			} else if(getMessageCode().equals(MESSAGE_TITLE_ADAPT_FINISH)){
				result = "document's title was adapted.";
			} else if(getMessageCode().equals(MESSAGE_ABSTRACT_ADAPT_FINISH)){
				result = "document's abstract was adapted.";
			} else if(getMessageCode().equals(MESSAGE_ACKNOWLEDGMENT_ADAPT_FINISH)){
				result = "document's acknowledgment was adapted.";
			} else if(getMessageCode().equals(MESSAGE_AUTHORS_ADAPT_FINISH)) {
				result = "document's author was adapted.";
			} else if(getMessageCode().equals(MESSAGE_BIBLIOGRAPHY_HEADER_ADAPT_FINISH)){
				result = "document's bibliography head was adapted.";
			} else if(getMessageCode().equals(MESSAGE_CAPTION_ADAPT_FINISH)) {
				result = "the captions were adpated.";
			} else if(getMessageCode().equals(MESSAGE_FOOTNOTE_ADAPT_FINISH)) {
				result = "the footnote style were adapted";
			} else if(getMessageCode().equals(MESSAGE_KEYWORD_ADAPT_FINISH)){
				result = "document's keyword was adapted.";
			} else if(getMessageCode().equals(MESSAGE_HEADING1_ADAPT_FINISH)){
				result = "document's heading1 was adapted.";
			} else if(getMessageCode().equals(MESSAGE_HEADING2_ADAPT_FINISH)){
				result = "document's heading2 was adapted.";
			} else if(getMessageCode().equals(MESSAGE_HEADING3_ADAPT_FINISH)){
				result = "document's heading3 was adapted.";
			} else if(getMessageCode().equals(MESSAGE_HEADING4_ADAPT_FINISH)){
				result = "document's heading4 was adapted.";
			} else if(getMessageCode().equals(MESSAGE_IMAGE_ADAPT_FINISH)){
				result = "The width and aligment of images will be adapted.";
			} else if(getMessageCode().equals(MESSAGE_NEWLINE_LAYOUT_APPENDED)){
				result = "A line with layout style was inserted.";
			} else if(getMessageCode().equals(MESSAGE_NUMBERFILE_ADAPT_FINISH)){
				result = "The numbering file has been adapted.";
			} else if(getMessageCode().equals(MESSAGE_PAGELAYOUT_ADAPT_FINISH)){
				result = "The page layout was adapted.";
			} else if(getMessageCode().equals(MESSAGE_REFERENCE_ADPAT_FINISH)){
				result = "document's references were adapted.";
			} else if(getMessageCode().equals(MESSAGE_REFERENCEHEADER_ADPAT_FINISH)){
				result = "document's reference header was adapted.";
			} else if(getMessageCode().equals(MESSAGE_STYLECLASS_ADAPT_FINISH)){
				result = "the style class file of the document was adapted.";
			} else if(getMessageCode().equals(MESSAGE_TABLECAPTION_ADAPT_FINISH)){
				result = "the table's captions was adapted.";
			} else if(getMessageCode().equals(MESSAGE_UNABLESAVE_ERROR)) {
				result = "the document cannot be saved.";
			} else if(getMessageCode().equals(MESSAGE_UNABLE_CREATE_NEW_STYLE_ITEM)) {
				result = "unable to create the new style for doc.";
			} else if(getMessageCode().equals(MESSAGE_UNABLE_FOUND_STYLE_BY_CLZID)) {
				result = "unable to find the style by class id";
			} else if(getMessageCode().equals(CHKRESULT_NULL)){
				result = "last style is null";
			} else if(getMessageCode().equals(MESSAGE_GENERAL_OK)) {
				result = "mission complete";
			} else {
				result = "message not defined";
			}
		} else if(getMessageDetails()!=null){
			result = getMessageDetails();
		} else {
			result = "unknow message";
		}
		return result;
	}
}
