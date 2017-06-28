package kkr.travel.components.travel.reader.excelpoi;

import java.io.File;

import kkr.common.errors.ConfigurationException;

public abstract class TravelReaderExcelPoiFwk {
	private boolean configured;

	protected File file;
	protected String sheet;

	public void config() throws ConfigurationException {
		configured = false;
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

	public String getSheet() {
		return sheet;
	}

	public void setSheet(String sheet) {
		this.sheet = sheet;
	}
}
