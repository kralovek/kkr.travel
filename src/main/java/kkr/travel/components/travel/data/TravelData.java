package kkr.travel.components.travel.data;

import java.net.URL;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;
import java.util.GregorianCalendar;

public class TravelData {

	private Integer year;
	private String name;
	private Date dateFrom;
	private Date dateTo;

	private URL immageTitle;
	private URL immageMap;

	private Collection<String> people = new ArrayList<String>();
	private Collection<String> places = new ArrayList<String>();

	private Collection<LinkData> traces = new ArrayList<LinkData>();
	private Collection<LinkData> journals = new ArrayList<LinkData>();
	private Collection<LinkData> photos = new ArrayList<LinkData>();

	public TravelData(String name) {
		this.name = name;
	}

	public Date getDateFrom() {
		return dateFrom;
	}

	public void setDateFrom(Date dateFrom) {
		this.dateFrom = dateFrom;
		Calendar calendar = new GregorianCalendar();
		calendar.setTime(dateFrom);
		year = calendar.get(Calendar.YEAR);
	}

	public Date getDateTo() {
		return dateTo;
	}

	public void setDateTo(Date dateTo) {
		this.dateTo = dateTo;
	}

	public URL getImmageTitle() {
		return immageTitle;
	}

	public void setImmageTitle(URL immageTitle) {
		this.immageTitle = immageTitle;
	}

	public URL getImmageMap() {
		return immageMap;
	}

	public void setImmageMap(URL immageMap) {
		this.immageMap = immageMap;
	}

	public Collection<String> getPeople() {
		return people;
	}

	public Collection<String> getPlaces() {
		return places;
	}

	public Collection<LinkData> getTraces() {
		return traces;
	}

	public Collection<LinkData> getJournals() {
		return journals;
	}

	public Collection<LinkData> getPhotos() {
		return photos;
	}

	public String getName() {
		return name;
	}
	
	public int getYear() {
		if (year == null) {
			throw new IllegalStateException("Object is not properly initialized");
		}
		return year;
	}
}
