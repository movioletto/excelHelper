package it.movioletto.helper.excel.enums;

public enum Operazione {
	NESSUNA("NESSUNA"),
	SOMMA("SOMMA"),
	MEDIA("MEDIA");

	private String codice;

	public String getCodice() {
		return codice;
	}

	Operazione(String codice) {
		this.codice = codice;
	}
}
