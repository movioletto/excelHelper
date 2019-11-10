package it.movioletto.helper.excel;

import it.movioletto.helper.excel.model.Foglio;
import it.movioletto.helper.excel.model.Tabella;
import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

import java.io.*;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * Rigourous Test :-)
     */
    public void testApp()
    {
        assertTrue( true );
    }

    /**
     * Write.createByList
     */
    public void testSimpleExcel()
    {
        List<Tabella> lista = new ArrayList<>();

        for(int i = 0;i < 10;i++) {
            Tabella tab = new Tabella();

            tab.setId(i);
            tab.setNome("Campo " + i);
            tab.setPrezzo((float) (i + 0.5));
            tab.setPrezzoIvato(new Double(tab.getPrezzo()) + ((tab.getPrezzo() * 22) / 100));
            tab.setPrezzoTotale(new BigDecimal(tab.getPrezzo() + tab.getPrezzoIvato()));

            Calendar c = Calendar.getInstance();
            c.setTime(new Date());

            c.add(Calendar.DATE, i);

            tab.setData(c.getTime());

            lista.add(tab);
        }

        try {
            byte[] doc = Write.createByList(lista);

            FileOutputStream fop = new FileOutputStream(new File("testSimpleExcel.xlsx"));

            fop.write(doc);
            fop.flush();
            fop.close();

            assertTrue( true );
        } catch (IOException | IllegalAccessException e) {
            e.printStackTrace();
            fail();
        }

    }

    /**
     * Rigourous Test :-)
     */
    public void testFoglioExcel()
    {
        Foglio foglio = new Foglio();

        List<Tabella> tabella0 = new ArrayList<>();

        for(int i = 0;i < 10;i++) {
            Tabella tab = new Tabella();

            tab.setId(i);
            tab.setNome("Campo t0" + i);
            tab.setPrezzo((float) (i + 0.5));
            tab.setPrezzoIvato(new Double(tab.getPrezzo()) + ((tab.getPrezzo() * 22) / 100));
            tab.setPrezzoTotale(new BigDecimal(tab.getPrezzo() + tab.getPrezzoIvato()));

            Calendar c = Calendar.getInstance();
            c.setTime(new Date());

            c.add(Calendar.DATE, i);

            tab.setData(c.getTime());

            tabella0.add(tab);
        }

        List<Tabella> tabella1 = new ArrayList<>();

        for(int i = 10;i < 20;i++) {
            Tabella tab = new Tabella();

            tab.setId(i);
            tab.setNome("Campo t2" + i);
            tab.setPrezzo((float) (i + 0.5));
            tab.setPrezzoIvato(new Double(tab.getPrezzo()) + ((tab.getPrezzo() * 22) / 100));
            tab.setPrezzoTotale(new BigDecimal(tab.getPrezzo() + tab.getPrezzoIvato()));

            Calendar c = Calendar.getInstance();
            c.setTime(new Date());

            c.add(Calendar.DATE, i);

            tab.setData(c.getTime());

            tabella1.add(tab);
        }

        foglio.setTabella0(tabella0);
        foglio.setTabella1(tabella1);

        try {
            byte[] doc = Write.createSheet(foglio);

            FileOutputStream fop = new FileOutputStream(new File("testFoglioExcel.xlsx"));

            fop.write(doc);
            fop.flush();
            fop.close();

            assertTrue( true );
        } catch (IOException | IllegalAccessException e) {
            e.printStackTrace();
            fail();
        }

    }

}
