package lib.ooxml.author;

import java.util.ArrayList;
import java.util.List;

public class Authors {

	private List<Author> authors;

	public List<Author> getAuthors() {
		return authors;
	}

	public void setAuthors(List<Author> authors) {
		this.authors = authors;
	}
	
	public String[] getStrings(){
		String[] result = new String[8];
		
		for(int index = 0; index < authors.size(); index++){
			Author author = authors.get(index);
			boolean isLastOne = (index+1)>=authors.size();
			result[0] += author.getName()+(isLastOne?"":", ");
			result[1] += author.getAffilication()+(isLastOne?"":", ");
			result[2] += author.getAddress()+(isLastOne?"":", ");
			result[3] += author.getCity()+(isLastOne?"":", ");
			result[4] += author.getZip()+(isLastOne?"":", ");
			result[5] += author.getCountry()+(isLastOne?"":", ");
			result[6] += author.getEmail()+(isLastOne?"":", ");
			result[7] += "";
		}
		
		return result;
	}
	public String[] getInfoStrings(){
		String[] result = new String[authors.size()];
		
		for(int index = 0; index < authors.size(); index++){
			Author author = authors.get(index);
			result[index] = author.getInfo();
		}
		return result;
	}
	public String getNameString(){
		String result = "";
		for(int index = 0; index < authors.size(); index++){
			boolean isLastOne = (index+1)>=authors.size();
			result += authors.get(index).getName()+(isLastOne?"":", ");
		}
		return result;
	}
	
	public Authors(){
		authors = new ArrayList<Author>();
	}
	
}
