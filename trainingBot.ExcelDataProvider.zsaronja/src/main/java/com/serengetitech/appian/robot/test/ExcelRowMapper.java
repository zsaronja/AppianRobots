package com.serengetitech.appian.robot.test;

import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.data.provider.api.IRowMapper;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellReference;

/**
 * Class to map an Excel file to the POJO class representing the object
 * {@link com.serengetitech.appian.robot.test.ExcelRow} and vice versa.
 * 
 * @author Jidoka
 */
public class ExcelRowMapper implements IRowMapper<IExcel, ExcelRow> {

	private static final int tvrtka = 0;
	private static final int oib = 1;
	private static final int internaOznaka = 2;
	private static final int ime = 3;
	private static final int prezime = 4;
	private static final int brRadneKnjizice = 5;
	private static final int jmbg = 6;
	private static final int datumRodjenja = 7;
	private static final int drzavljanstvo = 8;
	private static final int banka = 9;
	private static final int vrstaRacuna = 10;
	private static final int brojRacuna = 11;
	private static final int adresaPrebivalista = 12;
	private static final int adresaBoravista = 13;
	private static final int putniTroskovi = 14;
	private static final int email = 15;
	private static final int osnovnaPlaca = 16;
	private static final int mjestoTroska = 17;
	private static final int sifraMjestaTroska = 18;
	private static final int pkSifra = 19;
	private static final int pkGrad = 20;
	private static final int osobniOdbitak = 21;
	private static final int strucnaSprema = 22;
	private static final int zanimanje = 23;
	private static final int datumPrvogZaposlenja = 24;
	private static final int prvoZaposlenje = 25;
	private static final int vrstaUgovora = 26;
	private static final int pozicija = 27;
	private static final int posao = 28;
	private static final int lokacija = 29;
	private static final int brSatiTjedno = 30;
	
	/**
	 * @see IRowMapper#map(Object, int)
	 */
	@Override
	public ExcelRow map(IExcel data, int rowNum) {

		ExcelRow r = new ExcelRow();

		r.setTvrtka(data.getCellValueAsString(rowNum, tvrtka));
		r.setOib(data.getCellValueAsString(rowNum, oib));
		r.setInternaOznaka(data.getCellValueAsString(rowNum, internaOznaka));
		r.setIme(data.getCellValueAsString(rowNum, ime));
		r.setPrezime(data.getCellValueAsString(rowNum, prezime));
		r.setBrRadneKnjizice(data.getCellValueAsString(rowNum, brRadneKnjizice));
		r.setJmbg(data.getCellValueAsString(rowNum, jmbg));
		r.setDatumRodjenja(data.getCellValueAsString(rowNum, datumRodjenja));
		r.setDrzavljanstvo(data.getCellValueAsString(rowNum, drzavljanstvo));
		r.setBanka(data.getCellValueAsString(rowNum, banka));
		r.setVrstaRacuna(data.getCellValueAsString(rowNum, vrstaRacuna));
		r.setBrojRacuna(data.getCellValueAsString(rowNum, brojRacuna));
		r.setAdresaPrebivalista(data.getCellValueAsString(rowNum, adresaPrebivalista));
		r.setAdresaBoravista(data.getCellValueAsString(rowNum, adresaBoravista));
		r.setPutniTroskovi(data.getCellValueAsString(rowNum, putniTroskovi));
		r.setEmail(data.getCellValueAsString(rowNum, email));
		r.setOsnovnaPlaca(data.getCellValueAsString(rowNum, osnovnaPlaca));
		r.setMjestoTroska(data.getCellValueAsString(rowNum, mjestoTroska));
		r.setSifraMjestaTroska(data.getCellValueAsString(rowNum, sifraMjestaTroska));
		r.setPkSifra(data.getCellValueAsString(rowNum, pkSifra));
		r.setPkGrad(data.getCellValueAsString(rowNum, pkGrad));
		r.setOsobniOdbitak(data.getCellValueAsString(rowNum, osobniOdbitak));
		r.setStrucnaSprema(data.getCellValueAsString(rowNum, strucnaSprema));
		r.setZanimanje(data.getCellValueAsString(rowNum, zanimanje));
		r.setDatumPrvogZaposlenja(data.getCellValueAsString(rowNum, datumPrvogZaposlenja));
		r.setPrvoZaposlenje(data.getCellValueAsString(rowNum, prvoZaposlenje));
		r.setVrstaUgovora(data.getCellValueAsString(rowNum, vrstaUgovora));
		r.setPozicija(data.getCellValueAsString(rowNum, pozicija));
		r.setPosao(data.getCellValueAsString(rowNum, posao));
		r.setLokacija(data.getCellValueAsString(rowNum, lokacija));
		r.setBrSatiTjedno(data.getCellValueAsString(rowNum, brSatiTjedno));
		
		return isLastRow(r) ? null : r;
	}

	/**
	 * @see IRowMapper#update(Object, int, Object)
	 */
	@Override
	public void update(IExcel data, int rowNum, ExcelRow rowData) {

		data.setCellValueByRef(new CellReference(rowNum,  tvrtka), rowData.getTvrtka ());
		data.setCellValueByRef(new CellReference(rowNum,  oib), rowData.getOib ());
		data.setCellValueByRef(new CellReference(rowNum,  internaOznaka), rowData.getInternaOznaka ());
		data.setCellValueByRef(new CellReference(rowNum,  ime), rowData.getIme ());
		data.setCellValueByRef(new CellReference(rowNum,  prezime), rowData.getPrezime ());
		data.setCellValueByRef(new CellReference(rowNum,  brRadneKnjizice), rowData.getBrRadneKnjizice ());
		data.setCellValueByRef(new CellReference(rowNum,  jmbg), rowData.getJmbg ());
		data.setCellValueByRef(new CellReference(rowNum,  datumRodjenja), rowData.getDatumRodjenja ());
		data.setCellValueByRef(new CellReference(rowNum,  drzavljanstvo), rowData.getDrzavljanstvo ());
		data.setCellValueByRef(new CellReference(rowNum,  banka), rowData.getBanka ());
		data.setCellValueByRef(new CellReference(rowNum,  vrstaRacuna), rowData.getVrstaRacuna ());
		data.setCellValueByRef(new CellReference(rowNum,  brojRacuna), rowData.getBrojRacuna ());
		data.setCellValueByRef(new CellReference(rowNum,  adresaPrebivalista), rowData.getAdresaPrebivalista ());
		data.setCellValueByRef(new CellReference(rowNum,  adresaBoravista), rowData.getAdresaBoravista ());
		data.setCellValueByRef(new CellReference(rowNum,  putniTroskovi), rowData.getPutniTroskovi ());
		data.setCellValueByRef(new CellReference(rowNum,  email), rowData.getEmail ());
		data.setCellValueByRef(new CellReference(rowNum,  osnovnaPlaca), rowData.getOsnovnaPlaca ());
		data.setCellValueByRef(new CellReference(rowNum,  mjestoTroska), rowData.getMjestoTroska ());
		data.setCellValueByRef(new CellReference(rowNum,  sifraMjestaTroska), rowData.getSifraMjestaTroska ());
		data.setCellValueByRef(new CellReference(rowNum,  pkSifra), rowData.getPkSifra ());
		data.setCellValueByRef(new CellReference(rowNum,  pkGrad), rowData.getPkGrad ());
		data.setCellValueByRef(new CellReference(rowNum,  osobniOdbitak), rowData.getOsobniOdbitak ());
		data.setCellValueByRef(new CellReference(rowNum,  strucnaSprema), rowData.getStrucnaSprema ());
		data.setCellValueByRef(new CellReference(rowNum,  zanimanje), rowData.getZanimanje ());
		data.setCellValueByRef(new CellReference(rowNum,  datumPrvogZaposlenja), rowData.getDatumPrvogZaposlenja ());
		data.setCellValueByRef(new CellReference(rowNum,  prvoZaposlenje), rowData.getPrvoZaposlenje ());
		data.setCellValueByRef(new CellReference(rowNum,  vrstaUgovora), rowData.getVrstaUgovora ());
		data.setCellValueByRef(new CellReference(rowNum,  pozicija), rowData.getPozicija ());
		data.setCellValueByRef(new CellReference(rowNum,  posao), rowData.getPosao ());
		data.setCellValueByRef(new CellReference(rowNum,  lokacija), rowData.getLokacija ());
		data.setCellValueByRef(new CellReference(rowNum,  brSatiTjedno), rowData.getBrSatiTjedno ());

		// Auto size the column
		for (int i = 1; i < 32; i++) {
			data.getSheet().autoSizeColumn(i);
		}
	}

	/**
	 * The last row is determined when the first row without content in the first
	 * column is detected.
	 * <p>
	 * Another possibility could be to check also the second and the third columns.
	 * 
	 * @see IRowMapper#isLastRow(Object)
	 */
	@Override
	public boolean isLastRow(ExcelRow instance) {
		return instance == null || StringUtils.isBlank(instance.getTvrtka());
	}

	@Override
	public boolean mustBeProcessed(ExcelRow instance) {
		return IRowMapper.super.mustBeProcessed(instance);
	}
}
