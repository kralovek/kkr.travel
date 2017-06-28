package kkr.travel.components.travel.writer.html;

import java.io.File;

import kkr.common.errors.ConfigurationException;

public abstract class TravelWriterHtmlFwk {

	private boolean configured;

	protected File file;

	public void config() throws ConfigurationException {
		configured = false;

		if (file == null) {
			throw new ConfigurationException("Parameter 'file' is not configured");
		}

		configured = true;
	}

	public void testConfigured() {
		if (!configured) {
			throw new IllegalStateException(this.getClass().getName() + ": The component is not configured");
		}
	}

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}
}
