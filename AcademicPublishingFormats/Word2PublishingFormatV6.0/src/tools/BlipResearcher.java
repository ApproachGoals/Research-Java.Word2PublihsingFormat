package tools;

import java.util.ArrayList;
import java.util.List;

import org.docx4j.TraversalUtil.CallbackImpl;
import org.docx4j.dml.CTBlip;

public class BlipResearcher extends CallbackImpl {

	private List<CTBlip> blipList = new ArrayList<CTBlip>();
	@Override
	public List<Object> apply(Object arg0) {
		// TODO Auto-generated method stub
		
		if(arg0 instanceof CTBlip) {
			this.blipList.add((CTBlip)arg0);
		}
		
		return null;
	}
	
	public List<CTBlip> getBlipList() {
		return this.blipList;
	}

}
