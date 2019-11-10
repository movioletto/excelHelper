package it.movioletto.helper.excel.model;

import it.movioletto.helper.excel.annotation.ExcelCell;

import java.math.BigDecimal;
import java.util.Date;

public class Tabella {

	@ExcelCell(intestazione = "ID")
	private Integer id;

	@ExcelCell(intestazione = "Nome")
	private String nome;

	@ExcelCell(intestazione = "Prezzo")
	private Float prezzo;

	private Double prezzoIvato;

	@ExcelCell(intestazione = "Totale")
	private BigDecimal prezzoTotale;

	@ExcelCell(intestazione = "Data")
	private Date data;

	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getNome() {
		return nome;
	}

	public void setNome(String nome) {
		this.nome = nome;
	}

	public Float getPrezzo() {
		return prezzo;
	}

	public void setPrezzo(Float prezzo) {
		this.prezzo = prezzo;
	}

	public Double getPrezzoIvato() {
		return prezzoIvato;
	}

	public void setPrezzoIvato(Double prezzoIvato) {
		this.prezzoIvato = prezzoIvato;
	}

	public BigDecimal getPrezzoTotale() {
		return prezzoTotale;
	}

	public void setPrezzoTotale(BigDecimal prezzoTotale) {
		this.prezzoTotale = prezzoTotale;
	}

	public Date getData() {
		return data;
	}

	public void setData(Date data) {
		this.data = data;
	}
}
