package db;

public class Test {

	public static void main(String[] args){
		AcademicStylePool asp = new AcademicStylePool();
		asp.loadDefinition();
		if(asp.getStyles()!=null && asp.getStyles().get("acm")!=null && asp.getStyles().get("acm").getStyleMap()!=null){
			System.out.println( asp.getStyles().get("acm").getStyleMap().get(AcademicFormatStructureDefinition.TITLE).getStyleId());
			System.out.println( asp.getStyles().get("acm").getStyleMap().get(AcademicFormatStructureDefinition.TITLE).getType());
			System.out.println(asp.getStyles().get("acm").getStyleMap().get(AcademicFormatStructureDefinition.PAGEFORMAT).getPPr().getSectPr().getCols().getNum());
			System.out.println(asp.getStyles().get("acm").getStyleMap().get(AcademicFormatStructureDefinition.HEADING2).getPPr().getNumPr().getIlvl().getVal());
		}
		if(asp.getStyles()!=null && asp.getStyles().get("springer")!=null && asp.getStyles().get("springer").getStyleMap()!=null){
			System.out.println(asp.getStyles().get("springer").getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACT).getPPr().getRPr().getSz().getVal().intValue());				
		}
		if(asp.getStyles()!=null && asp.getStyles().get("elsevier")!=null && asp.getStyles().get("elsevier").getStyleMap()!=null){
			if(asp.getStyles().get("elsevier").getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACTHEADER).getPPr().getPBdr().getTop()!=null){
				System.out.println(asp.getStyles().get("elsevier").getStyleMap().get(AcademicFormatStructureDefinition.ABSTRACTHEADER).getPPr().getPBdr().getTop().getSz().intValue());				
			} else {
				System.out.println("error");
			}
		}
	}
}
