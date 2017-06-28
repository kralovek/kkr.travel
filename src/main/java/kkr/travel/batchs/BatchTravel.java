package kkr.travel.batchs;

import java.util.Collection;

import org.apache.log4j.Logger;

import kkr.common.errors.BaseException;
import kkr.travel.components.travel.data.TravelData;

public class BatchTravel extends BatchTravelFwk {
	private static final Logger LOG = Logger.getLogger(BatchTravel.class);

	public void run() throws BaseException {
		LOG.trace("BEGIN");
		try {
			Collection<TravelData> travelData = travelReader.readTravelData();
			
			travelWriter.writeTravelData(travelData);
			
			LOG.trace("OK");
		} finally {
			LOG.trace("END");
		}
	}
}
