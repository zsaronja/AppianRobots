package com.appian.robot.core.template;

import java.io.IOException;

import com.novayre.jidoka.client.api.exceptions.JidokaQueueException;
import com.novayre.jidoka.client.api.queue.AssignQueueParameters;
import com.novayre.jidoka.client.api.queue.IQueue;
import com.novayre.jidoka.client.api.queue.IQueueItem;
import com.novayre.jidoka.client.api.queue.IQueueManager;
import com.novayre.jidoka.client.api.queue.ReserveItemParameters;

/**
 * The Class QueueCommons.
 */
public class QueueCommons {

	/** The qmanager. */
	private IQueueManager qmanager;
	
	/** The reserve items parameters. */
	private ReserveItemParameters reserveItemsParameters;
	
	/**
	 * Inits the object.
	 *
	 * @param qmanager the qmanager
	 */
	public void init(IQueueManager qmanager) {
		
		this.qmanager = qmanager;
		
		reserveItemsParameters = new ReserveItemParameters();
		reserveItemsParameters.setUseOnlyCurrentQueue(true);
	}
	
	/**
	 * Gets the queue from id.
	 *
	 * @param queueId the queue id
	 * @return the queue from id
	 * @throws JidokaQueueException the jidoka queue exception
	 */
	public IQueue getQueueFromId(String queueId) throws JidokaQueueException {
		
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
	public IQueueItem getNextItem(IQueue currentQueue) throws JidokaQueueException {

		try {
			if (currentQueue == null) {
				return null;
			}

			return qmanager.reserveItem(reserveItemsParameters);

		} catch (JidokaQueueException e) {
			throw e;
		} catch (Exception e) {
			throw new JidokaQueueException(e);
		}
	}

	/**
	 * Gets the qmanager.
	 *
	 * @return the qmanager
	 */
	public IQueueManager getQmanager() {
		return qmanager;
	}

	/**
	 * Sets the qmanager.
	 *
	 * @param qmanager the qmanager to set
	 */
	public void setQmanager(IQueueManager qmanager) {
		this.qmanager = qmanager;
	}

	/**
	 * Gets the reserve items parameters.
	 *
	 * @return the reserveItemsParameters
	 */
	public ReserveItemParameters getReserveItemsParameters() {
		return reserveItemsParameters;
	}

	/**
	 * Sets the reserve items parameters.
	 *
	 * @param reserveItemsParameters the reserveItemsParameters to set
	 */
	public void setReserveItemsParameters(ReserveItemParameters reserveItemsParameters) {
		this.reserveItemsParameters = reserveItemsParameters;
	}

}
