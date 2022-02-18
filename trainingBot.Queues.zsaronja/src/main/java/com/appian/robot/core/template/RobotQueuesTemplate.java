package com.appian.robot.core.template;

import java.io.IOException;
import java.util.*;

import com.novayre.jidoka.client.api.execution.LaunchOptions;
import com.novayre.jidoka.client.api.execution.LaunchResult;
import com.novayre.jidoka.client.api.queue.*;
import org.apache.commons.lang3.StringUtils;

import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.exceptions.JidokaFatalException;
import com.novayre.jidoka.client.api.exceptions.JidokaQueueException;

import static com.novayre.jidoka.client.api.JidokaFactory.getServer;

/**
 * Robot Queues Template
 *
 * @author jidoka
 */
@Robot
public class RobotQueuesTemplate implements IRobot {

    /**
     * Server.
     */
    private IJidokaServer<?> server;

    /**
     * The IQueueManager instance.
     */
    private IQueueManager qmanager;

    /**
     * The current queue.
     */
    private IQueue currentQueue;
    /**
     * The current item index.
     */
    private int currentItemIndex;
    /**
     * The current item queue.
     */
    private IQueueItem currentItemQueue;

    /**
     * Added for RPA tutorial - Queues
     */
    private String newQueueId;
    private String robotId = "6201a73de4b0cc9a1990a638";

    /**
     * Action "startUp".
     * <p>
     * This method is used to initialize the Jidoka modules instances
     *
     * @throws Exception the exception
     */
    public boolean startUp() throws Exception {
        server = JidokaFactory.getServer();
        server.trace("TRACE startUp()");
        // initialize the QueueManager
        qmanager = server.getQueueManager();
        return IRobot.super.startUp();
    }


    /**
     * Action "start".
     */
    public void start() throws Exception {
        server.trace("TRACE start()");
        server.info("Starting process - start method ?");
    }

    /**
     * Action "Select queue"
     */
    public void selectQueue() throws Exception {
        server.trace("TRACE selectQueue()");
        if (StringUtils.isBlank(qmanager.preselectedQueue())) {
            server.warn("No queue selected");
            throw new Exception("No queue selected");
        } else {
            String selectedQueueID = qmanager.preselectedQueue();
            server.info("Selected queue ID: " + selectedQueueID);
            currentQueue = getQueueFromId(selectedQueueID);

            if (currentQueue == null) {
                server.debug("Queue not found");
                return;
            }
            server.setNumberOfItems(currentQueue.pendingItems());
        }
    }

    /**
     * Checks for more items.
     *
     * @return the string
     * @throws Exception the exception
     */
    public String hasMoreItems() throws Exception {
        // retrieve the next item in the queue
        server.trace("TRACE hasMoreItems()");
        currentItemQueue = getNextItem(currentQueue);

        if (currentItemQueue != null) {
            // set the stats for the current item
            server.setCurrentItem(currentItemIndex++, currentItemQueue.key());
            return "yes";
        }
        return "no";
    }

    /**
     * Log functional data.
     */
    public void logFunctionalData() {
        server.trace("TRACE logFunctionalData()");
        server.info(currentItemQueue.functionalData());
    }

    /**
     * Update item queue. This method is a sample to show how to update the first
     * element on the functional data map by adding the text " - MODIFIED"
     *
     * @throws JidokaQueueException the Jidoka queue exception
     */
    public void updateItemQueue() throws JidokaQueueException {

        server.trace("TRACE updateItemQueue()");
        // override the functional data in the queue item
        Map<String, String> funcData = currentItemQueue.functionalData();

        String firstKey = funcData.keySet().iterator().next();

        try {

            funcData.put(firstKey, funcData.get(firstKey) + " - MODIFIED");
            // release the item. The queue item result will be the same
            // as the current Item
            ReleaseItemWithOptionalParameters rip = new ReleaseItemWithOptionalParameters();
            rip.functionalData(funcData);
            rip.setProcess(EQueueItemReleaseProcess.FINISHED_OK);

            // Is mandatory to set the current item result before releasing the queue item
            server.setCurrentItemResultToOK(currentItemQueue.key());
            qmanager.releaseItem(rip);

        } catch (Exception e) {
            throw new JidokaQueueException(e);
        }
    }

    /**
     * Close queue action
     *
     * @throws IOException          Signals that an I/O exception has occurred.
     * @throws JidokaQueueException the jidoka queue exception
     */
    public void closeQueue() {
        server.trace("TRACE closeQueue()");
        try {
            // First we reserve the queue (other robots can't reserve the queue at the same time)
            ReserveQueueParameters rqp = new ReserveQueueParameters();
            rqp.setQueueId(currentQueue.queueId());
            IReservedQueue reservedQueue = qmanager.reserveQueue(rqp);

            // Robot can't reserve the current queue
            if (reservedQueue == null) {
                server.debug("Can't reserve the queue with ID: " + currentQueue.queueId());
                return;
            }

            // check for pending items.
            int pendingItems = reservedQueue.queue().pendingItems();

            ReleaseQueueParameters releaseQP = new ReleaseQueueParameters();
            server.info(String.format("The queue has %s pending items", pendingItems));

            if (pendingItems > 0) {
                // release the queue without closing
                releaseQP.closed(false);
                server.info("Queue released without closing");
            } else {
                // release the queue closing it
                releaseQP.closed(true);
                server.info("Queue closed and released");
            }
            qmanager.releaseQueue(releaseQP);
        } catch (Exception e) {
            throw new JidokaFatalException(e.getMessage(), e);
        }
    }

    /**
     * Action "end".
     *
     * @throws Exception the exception
     */
    public void end() {
        server.trace("TRACE end()");
        server.info("Ending process");
    }

    /**
     * Clean up.
     *
     * @return the string[]
     * @throws Exception the exception
     * @see com.novayre.jidoka.client.api.IRobot#cleanUp()
     */
    @Override
    public String[] cleanUp() throws Exception {
        return IRobot.super.cleanUp();
    }


    /**
     * Gets the queue from id.
     *
     * @param queueId the queue id
     * @return the queue from id
     * @throws JidokaQueueException the jidoka queue exception
     */
    private IQueue getQueueFromId(String queueId) throws JidokaQueueException {
        try {
            AssignQueueParameters qqp = new AssignQueueParameters();
            qqp.queueId(queueId);

            return qmanager.assignQueue(qqp);

        } catch (IOException e) {
            throw new JidokaQueueException(e);
        }
    }

    /**
     * Gets the next item.
     *
     * @return the next item
     * @throws JidokaQueueException the jidoka queue exception
     */
    private IQueueItem getNextItem(IQueue currentQueue) throws JidokaQueueException {
        server.trace("TRACE getNextItem()");
        try {
            if (currentQueue == null) {
                return null;
            }
            ReserveItemParameters reserveItemsParameters = new ReserveItemParameters();
            reserveItemsParameters.setUseOnlyCurrentQueue(true);
            return qmanager.reserveItem(reserveItemsParameters);
        } catch (Exception e) {
            throw new JidokaQueueException(e);
        }
    }

    /**
     * Added for RPA tutorial - Queues
     */
    public void createAndFillQueue() {
        server.trace("TRACE createAndFillQueue()");
        try {
            CreateQueueParameters qParam = new CreateQueueParameters();
            qParam.setDescription("Queue created from queue template on " + new Date());
            qParam.setName("QueueTemplate");
            qParam.setPriority(EPriority.NORMAL);
            qParam.setAttemptsByDefault(1);
            newQueueId = qmanager.createQueue(qParam);
            server.info("New Queue ID: " + newQueueId);
            createItem(newQueueId, "1", "United States", new HashMap<String, String>() {{
                put("Country", "United States");
                put("Capital City", "Washington D.C.");
            }});
            createItem(newQueueId, "2", "United Kingdom", new HashMap<String, String>() {{
                put("Country", "United Kingdom");
                put("Capital City", "London");
            }});
            createItem(newQueueId, "3", "Italy", new HashMap<String, String>() {{
                put("Country", "Italy");
                put("Capital City", "Rome");
            }});
            createItem(newQueueId, "4", "Spain", new HashMap<String, String>() {{
                put("Country", "Spain");
                put("Capital City", "Madrid");
            }});
            createItem(newQueueId, "5", "Croatia", new HashMap<String, String>() {{
                put("Country", "Croatia");
                put("Capital City", "Zagreb");
            }});
        } catch (Exception e) {
            throw new JidokaFatalException("The queue could not be created and filled");
        }
    }

    /**
     * Added for RPA tutorial - Queues
     */
    private void createItem(String queueID, String reference, String key, Map<String, String>
            functionalData) throws IOException, JidokaQueueException {
        server.trace("TRACE createItem()");
        CreateItemParameters itemParameters = new CreateItemParameters();
        // Set the item parameters
        itemParameters.setKey(key);
        itemParameters.setPriority(EPriority.NORMAL);
        itemParameters.setQueueId(queueID);
        itemParameters.setReference(reference);
        itemParameters.setFunctionalData(functionalData);
        qmanager.createItem(itemParameters);
        server.debug(String.format("Added item to queue %s with id %s", itemParameters.getQueueId(),
                itemParameters.getKey()));
    }

    /**
     * Added for RPA tutorial - Queues
     */
    public void selectCreatedQueue() throws Exception {
        server.trace("TRACE selectCreatedQueue()");
        currentQueue = getQueueFromId(newQueueId);
        if (currentQueue == null) {
            server.debug("Queue not found");
            return;
        }
        server.setNumberOfItems(currentQueue.pendingItems());
    }

    /**
     * Added for RPA tutorial - Queues
     */
    public void launchQueueRobot() {
        server.trace("TRACE launchQueueRobot()");
        try {
            server.debug("***** Preparing launch options *****");
            LaunchOptions options = new LaunchOptions();
            options.setDescription("Pokretanje Queue workera");
            options.queueId(newQueueId);
            options.robotName(robotId);
            options.executionsToLaunch(1);
            options.mustBeHere(false);
            options.setTestingMode(server.getExecution(0).getCurrentExecution().isTesting());
            server.debug("***** options PREPARED *****");
            List<LaunchResult> launchResults;
            launchResults = server.getExecution(0).launchRobots(options);
            server.debug("*****  Launch results options *****");
            server.debug("launchResults.isEmpty: " + launchResults.isEmpty());
            Iterator<LaunchResult> iterator = launchResults.iterator();
            while(iterator.hasNext())
            {
                LaunchResult next = iterator.next();
                server.debug(next.isLaunched());
                server.debug(next.getExecutionId());
                server.debug(next.getExecutionNumber());
                server.debug(next.getExecutionString());
                server.debug(next.getReason());
                server.debug(next.getErrorMsg());
            };
//            if (!launchResults.get(0).isLaunched()) {
//                server.error(launchResults.get(0).getErrorMsg());
//                throw new JidokaFatalException("Error launching the robot queue");
//            }
            server.info("Queue robot launched correctly");
        } catch (IOException e) {
            server.error(e.getMessage(), e);
            throw new JidokaFatalException("Error launching the robot queue");
        }
    }

    public void init()throws Exception {
        server = getServer();
        server.trace("TRACE init()");
        qmanager = server.getQueueManager();
    }
}
