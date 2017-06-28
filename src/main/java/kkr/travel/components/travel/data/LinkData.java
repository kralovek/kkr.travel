package kkr.travel.components.travel.data;

import java.net.URL;

public class LinkData {

	private String name;
	private URL url;

	public LinkData(String name, URL url) {
		super();
		this.name = name;
		this.url = url;
	}

	public String getName() {
		return name;
	}

	public URL getUrl() {
		return url;
	}
}
