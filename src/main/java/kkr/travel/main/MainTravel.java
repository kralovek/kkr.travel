package kkr.travel.main;

import kkr.common.errors.BaseException;
import kkr.common.main.AbstractMain;
import kkr.common.main.Config;
import kkr.travel.batchs.BatchTravel;

import org.apache.log4j.Logger;

public class MainTravel extends AbstractMain {
	private static final Logger LOG = Logger.getLogger(MainTravel.class);

	private static final String ID_BEAN_DEFAULT = "batchTravel";
	private static final String CONFIG_DEFAULT = "spring-main-travel.xml";

	public static final void main(String[] args) throws BaseException {
		LOG.trace("BEGIN");
		try {
			MainTravel mainTravel = new MainTravel();
			mainTravel.run(args);
			LOG.trace("OK");
		} finally {
			LOG.trace("END");
		}
	}
	
	private void run(String[] args) throws BaseException {
		Config config = config(getClass(), Config.class, args);
		LOG.trace("BEGIN");
		try {
			BatchTravel batch = createBean(config, BatchTravel.class, CONFIG_DEFAULT, ID_BEAN_DEFAULT);
			batch.run();
			LOG.trace("OK");
		} finally {
			LOG.trace("END");
		}
	}
}
