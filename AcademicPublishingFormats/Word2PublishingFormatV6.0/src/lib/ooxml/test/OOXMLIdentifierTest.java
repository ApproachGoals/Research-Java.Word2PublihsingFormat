package lib.ooxml.test;

import static org.junit.Assert.*;

import java.math.BigInteger;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.junit.Before;
import org.junit.Test;

public class OOXMLIdentifierTest {

	P p = null;
	@Before
	public void setUp() throws Exception {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
	}

	@Test
	public void testIsParagraphHasSectionNumber() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		
	}

	@Test
	public void testIsTitle() {
//		fail("Not yet implemented");
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setB(Context.getWmlObjectFactory().createBooleanDefaultTrue());
		r.getRPr().getB().setVal(true);
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsSubtitle() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsAbstractHeader() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsKeywordHeader() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsAcknowledgmentHeader() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsReferenceHeader() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsHeadingProperites() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(24));
		assertEquals(24, r.getRPr().getSz().getVal().intValue());
	}

	@Test
	public void testGetHeadingLvl() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		p.getPPr().setNumPr(Context.getWmlObjectFactory().createPPrBaseNumPr());
		p.getPPr().getNumPr().setNumId(Context.getWmlObjectFactory().createPPrBaseNumPrNumId());
		p.getPPr().getNumPr().getNumId().setVal(BigInteger.valueOf(99));
		p.getPPr().getNumPr().setIlvl(Context.getWmlObjectFactory().createPPrBaseNumPrIlvl());
		p.getPPr().getNumPr().getIlvl().setVal(BigInteger.valueOf(1));
		p.getPPr().setRPr(Context.getWmlObjectFactory().createParaRPr());
		p.getPPr().getRPr().setB(Context.getWmlObjectFactory().createBooleanDefaultTrue());
		p.getPPr().getRPr().getB().setVal(true);
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(24));
		assertEquals(2, p.getPPr().getNumPr().getIlvl().getVal().intValue()+1);
	}

	@Test
	public void testIsFigureCaption() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsTableCaption() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

	@Test
	public void testIsAppendix() {
		p = Context.getWmlObjectFactory().createP();
		p.setPPr(Context.getWmlObjectFactory().createPPr());
		p.getPPr().setPStyle(Context.getWmlObjectFactory().createPPrBasePStyle());
		R r = Context.getWmlObjectFactory().createR();
		r.setRPr(Context.getWmlObjectFactory().createRPr());
		r.getRPr().setSz(Context.getWmlObjectFactory().createHpsMeasure());
		r.getRPr().getSz().setVal(BigInteger.valueOf(28));
	}

}
