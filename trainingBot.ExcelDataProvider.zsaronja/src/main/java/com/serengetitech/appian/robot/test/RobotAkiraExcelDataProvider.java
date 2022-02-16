package com.serengetitech.appian.robot.test;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.novayre.jidoka.client.api.*;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.exceptions.JidokaFatalException;
import com.novayre.jidoka.client.api.exceptions.JidokaItemException;
import com.novayre.jidoka.data.provider.api.IJidokaDataProvider;
import com.novayre.jidoka.data.provider.api.IJidokaExcelDataProvider;
import org.apache.commons.lang3.exception.ExceptionUtils;

import java.io.File;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.HashMap;

import static com.novayre.jidoka.client.api.JidokaFactory.*;

@Robot
public class RobotAkiraExcelDataProvider implements IRobot {

    // Instance of the API server module implementation
    private IJidokaServer<?> server;

    /** The data provider. */
    private IJidokaExcelDataProvider<ExcelRow> dataProvider;

    /** The current item. */
    private ExcelRow currentItem;
    private String excelFile;

    // We use the annotation @JidokaMethod to define a method that could be use as
    // low code in the workflow.
    // Every parameter accessible from the workflow must be defined with the
    // annotation @JidokaParameter.
    // The parameter name is the value you fill on the workflow.
    @JidokaMethod(name = "Get Employees", description = "Read Employee data from AKIRA Excel report.")
    public String getEmployees(
            @JidokaParameter(name = "Excel report to read from (path)") String excelFilePath) {
        // Using server instance write an info message
        server.info(String.format("Accessing file %s", excelFilePath));
        // Create a new LaunchOptions object from the API to configure the launch to
        // perform

        String employeeJSON;
        try {
            HashMap<String, String> employee = processItem();
            ObjectMapper objMapper = new ObjectMapper();
            employeeJSON = objMapper.writeValueAsString(employee);
            server.debug("employeeJSON: ".concat(employeeJSON));
        } catch (Exception e) {
            server.error(e.getMessage(), e);
            throw new JidokaFatalException(
                    String.format("There is an error processing %s. ", excelFilePath), e);
        }
        server.info("Excel report processed.");
        return employeeJSON;
    }

    public void init() throws Exception {
        server = getServer();
        dataProvider = IJidokaDataProvider.getInstance( this, IJidokaDataProvider.Provider.EXCEL);
    }

    /**
     * Initializes the data provider.
     *
     * Action "start".
     *
     * @throws Exception
     *             in case any exception is thrown during the initialization
     */
    public void start(String excelFilePath) throws Exception {

        server.info("Initializing Data Provider with file: " + excelFilePath);

        // Path (String) to the file containing the items to process
        excelFile = Paths.get(server.getCurrentDir(), excelFilePath).toString();
        server.debug("excelFile: ".concat(excelFile));
        // Initialization of the Data Provider module using the RowMapper implemented
        dataProvider.init(excelFile, null, 0, new ExcelRowMapper());

        // Set the number of items relying on the Data Provider module
        server.setNumberOfItems(dataProvider.count());
        server.debug("dataProvider.count(): ".concat(String.valueOf(dataProvider.count())).concat(" Items set"));
    }

    /**
     * Checks for more items.
     *
     * @return the string representing the wire name in the workflow to follow.
     */
    public String hasMoreItems() {
        // To get the next row, we rely again on the Data Provider module
        server.debug("dataProvider.nextRow(): ".concat(String.valueOf(dataProvider.nextRow())));
        return dataProvider.nextRow() ? "yes" : "no";
    }

    /**
     * Processes an item.
     *
     */
    public HashMap<String, String> processItem() {
//    public void processItem() {
        HashMap<String, String> retVal = new HashMap<>();

        // Get the current item through Data Provider
        currentItem = dataProvider.getCurrentItem();
        retVal.put("ime", currentItem.getIme());
        retVal.put("prezime", currentItem.getPrezime());
        retVal.put("email", currentItem.getEmail());
//        // The key to use is the literal "row" plus the number of the item
        String itemKey = "row " + dataProvider.getCurrentItemNumber();
        server.setCurrentItem(dataProvider.getCurrentItemNumber(), itemKey);
//        // The process is very simple: to concatenate the 3 columns to get the result
        String item = currentItem.getIme().concat(" ").concat(currentItem.getPrezime()).concat(" - ").concat(currentItem.getEmail());
//        // We consider this item is OK
        server.setCurrentItemResultToOK(item);
        return retVal;
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
     * It is a private method to be called during the cleanup
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
     * @return
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
     */
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
            return "hasMoreItems";
        }

        // Unknown exception. Throw it
        throw exception;
    }
}
