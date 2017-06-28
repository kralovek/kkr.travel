package kkr.travel.batchs;

import kkr.common.errors.ConfigurationException;
import kkr.travel.components.travel.reader.TravelReader;
import kkr.travel.components.travel.writer.TravelWriter;

public abstract class BatchTravelFwk {
	private boolean configured;

	protected TravelReader travelReader;
	protected TravelWriter travelWriter;

	public void config() throws ConfigurationException {
		configured = false;

		if (travelReader == null) {
			throw new ConfigurationException("Parameter 'travelReader' is not configured");
		}

		if (travelWriter == null) {
			throw new ConfigurationException("Parameter 'travelWriter' is not configured");
		}

		configured = true;
	}

	public void testConfigured() {
		if (!configured) {
			throw new IllegalStateException(this.getClass().getName() + ": The component is not configured");
		}
	}

	public TravelReader getTravelReader() {
		return travelReader;
	}

	public void setTravelReader(TravelReader travelReader) {
		this.travelReader = travelReader;
	}

	public TravelWriter getTravelWriter() {
		return travelWriter;
	}

	public void setTravelWriter(TravelWriter travelWriter) {
		this.travelWriter = travelWriter;
	}
}
