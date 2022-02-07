package com.appian.robot.core.template;

import java.io.IOException;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;

import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.exceptions.JidokaFatalException;
import com.novayre.jidoka.client.api.exceptions.JidokaQueueException;
import com.novayre.jidoka.client.api.queue.AssignQueueParameters;
import com.novayre.jidoka.client.api.queue.EQueueItemReleaseProcess;
import com.novayre.jidoka.client.api.queue.IQueue;
import com.novayre.jidoka.client.api.queue.IQueueItem;
import com.novayre.jidoka.client.api.queue.IQueueManager;
import com.novayre.jidoka.client.api.queue.IReservedQueue;
import com.novayre.jidoka.client.api.queue.ReleaseItemWithOptionalParameters;
import com.novayre.jidoka.client.api.queue.ReleaseQueueParameters;
import com.novayre.jidoka.client.api.queue.ReserveItemParameters;
import com.novayre.jidoka.client.api.queue.ReserveQueueParameters;

/**
 * Robot Queues Template
 *
 * @author jidoka
 */
@Robot
public class RobotQueuesTemplate implements IRobot {

	/** Server. */
	private IJidokaServer<?> server;

	/** The IQueueManager instance. */
	private IQueueManager qmanager;

	/** The current queue. */
	private IQueue currentQueue;
	/** The current item index. */
	private int currentItemIndex;
	/** The current item queue. */
	private IQueueItem currentItemQueue;

	/**
	 * Action "startUp".
	 * 
	 * This method is used to initalize the Jidoka modules instances
	 *
	 * @throws Exception the exception
	 */
	public boolean startUp() throws Exception {
		server = (IJidokaServer<?>) JidokaFactory.getServer();
		// intialize the QueueManager
		qmanager = server.getQueueManager();
		return IRobot.super.startUp();
	}

	
	/**
	 * Action "start".
	 * 
	 */
	public void start() throws Exception {
		server.info("Starting process");
	}

	/**
	 * Action "Select queue"
	 *
	 */
	public void selectQueue() throws Exception {

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
		server.info(currentItemQueue.functionalData());
	}

	/**
	 * Update item queue. This method is a sample to show how toupdate the first 
	 * element on the functional data map by adding the text " - MODIFIED"
	 *
	 * @throws JidokaQueueException the Jidoka queue exception
	 */
	public void updateItemQueue() throws JidokaQueueException {

		// override the functional data in the queue item
		Map<String, String> funcData = currentItemQueue.functionalData();

		String firstKey = funcData.keySet().iterator().next();

		try {
			
			funcData.put(firstKey, funcData.get(firstKey) + " - MODIFIED");
			// release the item. The queue item result will be the same
			// as the currenItem
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
	 * @return the string
	 * @throws IOException          Signals that an I/O exception has occurred.
	 * @throws JidokaQueueException the jidoka queue exception
	 */
	public void closeQueue() {

		try {
			// First we reserve the queue (other robots can't reserve the queue at the same time)
			ReserveQueueParameters rqp = new ReserveQueueParameters();
			rqp.setQueueId(currentQueue.queueId());
			IReservedQueue reservedQueue = qmanager.reserveQueue(rqp);
			
			// Robot can't reserver the current queue
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

			IQueue queue = qmanager.assignQueue(qqp);
			return queue;
				
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
}
