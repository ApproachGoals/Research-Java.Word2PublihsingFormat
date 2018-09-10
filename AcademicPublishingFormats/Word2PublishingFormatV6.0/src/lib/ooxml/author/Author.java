package lib.ooxml.author;

public class Author {
	private String name;
	private String affilication;
	private String address;
	private String city;
	private String zip;
	private String country;
	private String email;
	
	private String info1;
	private String info2;
	private String info3;
	private String info4;
	public String getInfo1() {
		return info1;
	}
	public void setInfo1(String info1) {
		this.info1 = info1;
	}
	public String getInfo2() {
		return info2;
	}
	public void setInfo2(String info2) {
		this.info2 = info2;
	}
	public String getInfo3() {
		return info3;
	}
	public void setInfo3(String info3) {
		this.info3 = info3;
	}
	public String getInfo4() {
		return info4;
	}
	public void setInfo4(String info4) {
		this.info4 = info4;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getAffilication() {
		return affilication;
	}
	public void setAffilication(String affilication) {
		this.affilication = affilication;
	}
	public String getAddress() {
		return address;
	}
	public void setAddress(String address) {
		this.address = address;
	}
	public String getCity() {
		return city;
	}
	public void setCity(String city) {
		this.city = city;
	}
	public String getZip() {
		return zip;
	}
	public void setZip(String zip) {
		this.zip = zip;
	}
	public String getCountry() {
		return country;
	}
	public void setCountry(String country) {
		this.country = country;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getAffiInfoString(){
		String result = "";
		if(getAffilication()!=null && getAffilication().length()>0){
			result += getAffilication()+ ", ";
		}
		if(getAddress()!=null && getAddress().length()>0){
			result += getAddress()+ ", ";
		}
		if(getCity()!=null && getCity().length()>0){
			result += getCity()+ ", ";
		}
		if(getZip()!=null && getZip().length()>0){
			result += getZip()+ ", ";
		}
		if(getCountry()!=null && getCountry().length()>0){
			result += getCountry();
		}
		return result;
	}
	public String getAffiInfoString2(){
		String result = "";
		if(getAffilication()!=null && getAffilication().length()>0){
			result += getAffilication()+ ", ";
		}
		if(getInfo1()!=null && getInfo1().length()>0){
			result += getInfo1()+ ", ";
		}
		if(getInfo2()!=null && getInfo2().length()>0){
			result += getInfo2()+ ", ";
		}
		if(getInfo3()!=null && getInfo3().length()>0){
			result += getInfo3()+ ", ";
		}
		if(getInfo4()!=null && getInfo4().length()>0){
			result += getInfo4();
		}
		return result;
	}
	public void setInfo(int index, String info){
		switch (index) {
		case 1:
			setInfo1(info);
			break;
		case 2:
			setInfo2(info);
			break;
		case 3:
			setInfo3(info);
			break;
		case 4:
			setInfo4(info);
			break;
		default:
			break;
		}
	}
	public String getInfo(){
		String result = "";
		result += (getInfo1()!=null?getInfo1()+", ":"")+(getInfo2()!=null?getInfo2()+", ":"")+(getInfo3()!=null?getInfo3()+", ":"")+(getInfo4()!=null?getInfo4():"");
		if(result.endsWith(", ")){
			result = result.substring(0, result.length()-2);
		}
		return result;
	}
}
