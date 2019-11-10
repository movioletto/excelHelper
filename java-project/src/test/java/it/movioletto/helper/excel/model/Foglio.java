package it.movioletto.helper.excel.model;

import it.movioletto.helper.excel.annotation.ExcelSheet;
import it.movioletto.helper.excel.annotation.ExcelTable;

import java.util.List;

@ExcelSheet(titolo = "Foglio 0")
public class Foglio {

	@ExcelTable(titolo = "Tabella numero 0")
	private List<Tabella> tabella0;

	private List<Tabella> tabella1;

	public List<Tabella> getTabella0() {
		return tabella0;
	}

	public void setTabella0(List<Tabella> tabella0) {
		this.tabella0 = tabella0;
	}

	public List<Tabella> getTabella1() {
		return tabella1;
	}

	public void setTabella1(List<Tabella> tabella1) {
		this.tabella1 = tabella1;
	}
}
