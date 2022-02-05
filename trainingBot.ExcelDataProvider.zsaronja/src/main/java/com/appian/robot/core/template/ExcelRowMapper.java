package com.appian.robot.core.template;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.util.CellReference;

import com.novayre.jidoka.data.provider.api.IExcel;
import com.novayre.jidoka.data.provider.api.IRowMapper;

/**
 * Class to map an Excel file to the POJO class representing the object
 * {@link ExcelRow} and vice versa.
 * 
 * @author Jidoka
 */
public class ExcelRowMapper implements IRowMapper<IExcel, ExcelRow> {

	/**
	 * First column.
	 */
	private static final int COL_1 = 0;

	/**
	 * Second column.
	 */
	private static final int COL_2 = 1;

	/**
	 * Third column. 
	 */
	private static final int COL_3 = 2;

	/**
	 * Column with the result.
	 */
	private static final int RESULT_COL = 3;

	/** Added for RPA tutorial - Excel Data Provider */
	/**
	 * Column with the status.
	 */
	private static final int STATUS_COL = 4;

	/**
	 * @see com.novayre.jidoka.data.provider.api.IRowMapper#map(java.lang.Object, int)
	 */
	@Override
	public ExcelRow map(IExcel data, int rowNum) {

		ExcelRow r = new ExcelRow();

		r.setCol1(data.getCellValueAsString(rowNum, COL_1));
		r.setCol2(data.getCellValueAsString(rowNum, COL_2));
		r.setCol3(data.getCellValueAsString(rowNum, COL_3));
		r.setResult(data.getCellValueAsString(rowNum, RESULT_COL));

		/** Added for RPA tutorial - Excel Data Provider */
		r.setStatus(data.getCellValueAsString(rowNum, STATUS_COL));

		return isLastRow(r) ? null : r;
	}

	/**
	 * @see com.novayre.jidoka.data.provider.api.IRowMapper#update(java.lang.Object, int, java.lang.Object)
	 */
	@Override
	public void update(IExcel data, int rowNum, ExcelRow rowData) {

		data.setCellValueByRef(new CellReference(rowNum, COL_1), rowData.getCol1());
		data.setCellValueByRef(new CellReference(rowNum, COL_2), rowData.getCol2());
		data.setCellValueByRef(new CellReference(rowNum, COL_3), rowData.getCol3());
		data.setCellValueByRef(new CellReference(rowNum, RESULT_COL), rowData.getResult());

		/** Added for RPA tutorial - Excel Data Provider */
		data.setCellValueByRef(new CellReference(rowNum, STATUS_COL), rowData.getStatus());

		// Auto size the column
		for (int i = 1; i < 5; i++) {
			data.getSheet().autoSizeColumn(i);
		}
	}

	/**
	 * The last row is determined when the first row without content in the first
	 * column is detected.
	 * <p>
	 * Another possibility could be to check also the second and the third columns.
	 * 
	 * @see com.novayre.jidoka.data.provider.api.IRowMapper#isLastRow(java.lang.Object)
	 */
	@Override
	public boolean isLastRow(ExcelRow instance) {
		return instance == null || StringUtils.isBlank(instance.getCol1());
	}

	/** Added for RPA tutorial - Excel Data Provider */
	/**
	 * We are going to filter the rows where the first column has the value "R"
	 *
	 *
	 * @see com.novayre.jidoka.data.provider.api.IRowMapper#mustBeProcessed(java.lang.Object)
	 */
	@Override
	public boolean mustBeProcessed(ExcelRow instance) {
		return (instance!=null && instance.getCol1()!=null && !instance.getCol1().equals("R"));
	}
}
