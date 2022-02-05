package com.appian.robot.core.template;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;

import org.apache.commons.lang3.exception.ExceptionUtils;

import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.exceptions.JidokaItemException;
import com.novayre.jidoka.data.provider.api.IJidokaDataProvider;
import com.novayre.jidoka.data.provider.api.IJidokaDataProvider.Provider;
import com.novayre.jidoka.data.provider.api.IJidokaExcelDataProvider;

/**
 * Robot Data Provider Template
 *
 * @author Jidoka
 */
@Robot
public class RobotDataProviderExcelTemplate implements IRobot {
	
	/** The Constant EXCEL_FILENAME. */
	private static final String EXCEL_FILENAME = "robot-dataprovider-excel-inputfile.xlsx";

	/** Server. */
	private IJidokaServer<?> server;

	/** The data provider. */
	private IJidokaExcelDataProvider<ExcelRow> dataProvider;

	/** The current item. */
	private ExcelRow currentItem;

	/** The excel file. */
	private String excelFile;

	/**
	 * Action "startUp".
	 * <p>
	 * This method is overwritten to initialize the Appian RPA modules instances.
	 */
	@Override
	public boolean startUp() throws Exception {

		server = (IJidokaServer<?>) JidokaFactory.getServer();
		
		dataProvider = IJidokaDataProvider.getInstance(this, Provider.EXCEL);
		
		return IRobot.super.startUp();
	}
	
	/**
	 * Initializes the data provider.
 	 *
	 * Action "start".
	 *
	 * @throws Exception
	 *             in case any exception is thrown during the initialization
	 */
	public void start() throws Exception {
		
		server.info("Initializing Data Provider with file: " + EXCEL_FILENAME);
		
		// Path (String) to the file containing the items to process
		excelFile = Paths.get(server.getCurrentDir(), EXCEL_FILENAME).toString();
		
		// Initialization of the Data Provider module using the RowMapper implemented
		dataProvider.init(excelFile, null, 0, new ExcelRowMapper());
		
		// Set the number of items relying on the Data Provider moduleç
		server.setNumberOfItems(dataProvider.count());
	}


	/**
	 * Checks for more items.
	 *
	 * @return the string representing the wire name in the workflow to follow.
	 */
	public String hasMoreItems() {
		
		// To get the next row, we rely again on the Data Provider module
		return dataProvider.nextRow() ? "yes" : "no";
	}

	/**
	 * Processes an item.
	 * 
	 * In this template example, the processing consists of concatenating the first
	 * 3 columns to get the string result and update the last column.
	 */
	public void processItem() {

		// Get the current item through Data Provider
		currentItem = dataProvider.getCurrentItem();
		
		// The key to use is the literal "row" plus the number of the item
		String itemKey = "row " + dataProvider.getCurrentItemNumber();
		server.setCurrentItem(dataProvider.getCurrentItemNumber(), itemKey);
		
		// The process is very simple: to concatenate the 3 columns to get the result
		String result = currentItem.getCol1() + currentItem.getCol2() + currentItem.getCol3();
		currentItem.setResult(result);

		/** Added for RPA tutorial - Excel Data Provider */
		currentItem.setStatus("Processed");

		// Update the item in the Excel file through Data Provider
		dataProvider.updateItem(currentItem);
		
		// We consider this item is OK
		server.setCurrentItemResultToOK(currentItem.getResult());
	}
		
	/**
	 * Last action of the robot.
	 */
	public void end() {

		server.info("End process");
	}

	/**
	 * Method that closes the data provider.
	 * 
	 * It is a private method to be called during the clean up 
 	 * to assure the data provider is correctly closed.
	 * 
	 * @throws IOException
	 *             if an I/O error occurs
	 */
	private void closeDataProvider() throws IOException {
		
		if (dataProvider != null)  {
			
			server.info("Closing Data Provider...");
			
			dataProvider.close();
			dataProvider = null;
		}
	}

	/**
	 * Cleans up.
	 * 
	 * Besides returning the updated Excel file, it closes the data
	 * provider. 
	 *
	 * @return an array with the paths of the files to return to the console
	 * @throws Exception
	 *             in case any exception is thrown
	 * 
	 * @see com.novayre.jidoka.client.api.IRobot#cleanUp()
	 */
	@Override
	public String[] cleanUp() throws Exception {
		
		closeDataProvider();

		if (new File(excelFile).exists()) {
			return new String[] { excelFile };
		}

		return IRobot.super.cleanUp();
	}
	
	/**
	 * Manages exceptions that may arise during the robot execution.
	 * 
	 * @see com.novayre.jidoka.client.api.IRobot#manageException(java.lang.String,
	 *      java.lang.Exception)
	 */
	@Override
	public String manageException(String action, Exception exception) throws Exception {
		
		// Optionally, we send the exception to the execution log
		server.warn(exception.getMessage(), exception);
		
		/*
		 * We take advantage of the Apache Commons ExceptionUtils class to know if a
		 * specific exception was thrown throughout the code.
		 * Since we threw a JidokaItemException, it's the one to be searched in the
		 * exceptions stack trace. If found, the flow goes to the next item by telling
		 * the next method to execute is 'moreItems()'.
		 * If another exception is found, it is propagated, so the robot ends with a
		 * failure.
		 */
		if (ExceptionUtils.indexOfThrowable(exception, JidokaItemException.class) >= 0) {
			server.setCurrentItemResultToWarn("Exception processing the item!");

			/** Added for RPA tutorial - Excel Data Provider */
			currentItem.setStatus("Failed");
			dataProvider.updateItem(currentItem);

			return "hasMoreItems";
		}
		
		// Unknown exception. Throw it
		throw exception;
	}

}
