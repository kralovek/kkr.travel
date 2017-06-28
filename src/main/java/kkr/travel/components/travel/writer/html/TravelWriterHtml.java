package kkr.travel.components.travel.writer.html;

import java.io.IOException;
import java.io.PrintWriter;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import kkr.common.errors.BaseException;
import kkr.common.errors.TechnicalException;
import kkr.common.utils.UtilsResource;
import kkr.travel.components.travel.data.LinkData;
import kkr.travel.components.travel.data.TravelData;
import kkr.travel.components.travel.writer.TravelWriter;
import org.apache.log4j.Logger;

public class TravelWriterHtml extends TravelWriterHtmlFwk implements TravelWriter {
	private static final Logger LOG = Logger.getLogger(TravelWriterHtml.class);

	private static DateFormat DATE_FORMAT = new SimpleDateFormat("dd/MM/yyyy");

	private static Comparator<TravelData> COMPARATOR_TRAVEL_DATA = new Comparator<TravelData>() {

		public int compare(TravelData travelData1, TravelData travelData2) {
			return -(new Integer(travelData1.getYear()).compareTo(travelData2.getYear()));
		}
	};
	
	public void writeTravelData(Collection<TravelData> travelDatas) throws BaseException {
		LOG.trace("BEGIN");
		try {
			testConfigured();

			List<TravelData> sortedTravelDatas = new ArrayList<TravelData>(travelDatas);
			Collections.sort(sortedTravelDatas, COMPARATOR_TRAVEL_DATA); 
			
			PrintWriter printWriter = null;
			try {
				printWriter = new PrintWriter(file);

				writeTravelData(sortedTravelDatas, printWriter);

				printWriter.close();
				printWriter = null;
			} catch (IOException ex) {
				throw new TechnicalException("Cannot create the file: " + file.getAbsolutePath(), ex);
			} finally {
				UtilsResource.closeResource(printWriter);
			}
			LOG.trace("OK");
		} finally {
			LOG.trace("END");
		}
	}

	private void writeTravelData(Collection<TravelData> travelDatas, PrintWriter printWriter) throws BaseException {
		writeHeader(printWriter);

		Integer lastYear = null;
		for (TravelData travelData : travelDatas) {
			if (lastYear == null || lastYear != travelData.getYear()) {
				lastYear = travelData.getYear(); 
				writeYear(printWriter, lastYear);
			}
			writeItem(printWriter, travelData);
		}

		writeTail(printWriter);
	}

	private void writeHeader(PrintWriter printWriter) {
		printWriter.println("<HTML>");
		printWriter.println("<HEAD>");
		printWriter.println("<META CHARSET=\"UTF-8\" />");
		printWriter.println("<TITLE>Travel</TITLE>");
		printWriter.println("<BODY>");
		printWriter.println();
	}

	private void writeYear(PrintWriter printWriter, int year) {
		printWriter.println();
		printWriter.println("<DIV>");
		printWriter.println("<FONT SIZE=\"6\">" + year + "</FONT>");
		printWriter.println("</DIV>");
	}

	private void writeItem(PrintWriter printWriter, TravelData travelData) {
		printWriter.println();
		printWriter.println("<DIV>");
		printWriter.println("<TABLE VALIGN=\"top\">");
		printWriter.println("<TBODY>");
		printWriter.println("<TR>");
		printWriter.println("<!-- IMMAGE TITLE -->");
		printWriter.println("<TD VALIGN=\"top\">");
		if (travelData.getImmageTitle() != null) {
			URL url = findURL(travelData.getPhotos());
			if (url != null) {
				printWriter.println("<A HREF=\""+url.toString()+"\" TARGET=\"_blank\">");
			}
			printWriter.println("<IMG SRC=\"" + travelData.getImmageTitle().toString() + "\" WIDTH=\"200\" />");
			if (url != null) {
				printWriter.println("</A>");
			}
		}
		printWriter.println("</TD>");

		printWriter.println("<!-- DATA -->");
		printWriter.println("<TD WIDTH=\"100%\" VALIGN=\"top\">");
		printWriter.println("    <TABLE WIDTH=\"100%\">");
		printWriter.println("    <TBODY>");

		//
		// NAME
		//
		printWriter.println("");
		printWriter.println("    <TR>");
		printWriter.println("    <TD COLSPAN=\"2\">");
		printWriter.print("<B><FONT COLOR=\"#0b5394\" SIZE=\"4\">");
		printWriter.print(travelData.getName());
		printWriter.print("</FONT></B>");
		printWriter.print("</TD>");
		printWriter.println("    </TR>");

		//
		// DATE
		//
		if (travelData.getDateFrom() != null) {
			printWriter.println("");
			printWriter.println("    <TR>");
			printWriter.println("    <TD><B>Date:</B></TD>");
			printWriter.print("    <TD WIDTH=\"100%\">");
			printWriter.print(DATE_FORMAT.format(travelData.getDateFrom()));
			if (travelData.getDateTo() != null && !travelData.getDateFrom().equals(travelData.getDateTo())) {
				printWriter.println(" - " + DATE_FORMAT.format(travelData.getDateTo()));
			}
			printWriter.println("</TD>");
			printWriter.println("    </TR>");
		}

		//
		// PEOPLE
		//
		if (travelData.getPeople() != null && !travelData.getPeople().isEmpty()) {
			printWriter.println("");
			printWriter.println("    <TR>");
			printWriter.println("    <TD><B>People:</B></TD>");
			printWriter.print("    <TD WIDTH=\"100%\">");
			writeValueList(printWriter, travelData.getPeople());
			printWriter.println("</TD>");
			printWriter.println("    </TR>");
		}

		//
		// PLACES
		//
		if (travelData.getPlaces() != null && !travelData.getPlaces().isEmpty()) {
			printWriter.println("");
			printWriter.println("    <TR>");
			printWriter.println("    <TD><B>Places:</B></TD>");
			printWriter.print("    <TD WIDTH=\"100%\">");
			writeValueList(printWriter, travelData.getPlaces());
			printWriter.println("</TD>");
			printWriter.println("    </TR>");
		}

		//
		// PHOTOS
		//
		if (travelData.getPhotos() != null && !travelData.getPhotos().isEmpty()) {
			printWriter.println("");
			printWriter.println("    <TR>");
			printWriter.println("    <TD><B>Photos:</B></TD>");
			printWriter.print("    <TD WIDTH=\"100%\">");
			writeLinkList(printWriter, travelData.getPhotos());
			printWriter.println("</TD>");
			printWriter.println("    </TR>");
		}

		//
		// TRACES
		//
		if (travelData.getTraces() != null && !travelData.getTraces().isEmpty()) {
			printWriter.println("");
			printWriter.println("    <TR>");
			printWriter.println("    <TD><B>Traces:</B></TD>");
			printWriter.print("    <TD WIDTH=\"100%\">");
			writeLinkList(printWriter, travelData.getTraces());
			printWriter.println("</TD>");
			printWriter.println("    </TR>");
		}

		//
		// JOURNAL
		//
		if (travelData.getJournals() != null && !travelData.getJournals().isEmpty()) {
			printWriter.println("");
			printWriter.println("    <TR>");
			printWriter.println("    <TD><B>Journal:</B></TD>");
			printWriter.print("    <TD WIDTH=\"100%\">");
			writeLinkList(printWriter, travelData.getJournals());
			printWriter.println("</TD>");
			printWriter.println("    </TR>");
		}

		printWriter.println("");
		printWriter.println("    </TBODY>");
		printWriter.println("    </TABLE>");
		printWriter.println("</TD>");

		printWriter.println("<!-- IMMAGE MAP -->");
		printWriter.println("<TD VALIGN=\"top\">");
		if (travelData.getImmageMap() != null) {
			URL url = findURL(travelData.getTraces());
			if (url != null) {
				printWriter.println("<A HREF=\""+url.toString()+"\" TARGET=\"_blank\">");
			}
			printWriter.print("<IMG SRC=\"" + travelData.getImmageMap().toString() + "\" WIDTH=\"200\" />");
			if (url != null) {
				printWriter.println("</A>");
			}
		}
		printWriter.println("</TD>");

		printWriter.println("</TR>");
		printWriter.println("</TBODY>");
		printWriter.println("</TABLE>");
		printWriter.println("</DIV>");
	}

	private void writeTail(PrintWriter printWriter) {
		printWriter.println();
		printWriter.println("</BODY>");
		printWriter.println("</HTML>");
	}

	private void writeValueList(PrintWriter printWriter, Collection<String> values) {
		boolean first = true;
		for (String item : values) {
			if (!first) {
				printWriter.print(", ");
			}
			printWriter.print(item);
			first = false;
		}
	}

	private void writeLinkList(PrintWriter printWriter, Collection<LinkData> links) {
		boolean first = true;
		for (LinkData item : links) {
			if (isEmpty(item.getName()) && isEmpty(item.getUrl())) {
				continue;
			}
			if (!first) {
				printWriter.print(", ");
			}

			if (item.getUrl() != null) {
				printWriter.print("<A HREF=\"" + item.getUrl().toString() + "\" TARGET=\"_blank\">");
			}
			if (item.getName() != null) {
				printWriter.print(item.getName());
			} else {
				printWriter.print(item.getUrl().toString());
			}
			if (item.getUrl() != null) {
				printWriter.print("</A>");
			}
			first = false;
		}
	}

	private boolean isEmpty(Object value) {
		if (value == null) {
			return true;
		}

		if (value instanceof String) {
			return ((String) value).isEmpty();
		}

		return false;
	}
	
	private URL findURL(Collection<LinkData> links) {
		if (links == null || links.isEmpty()) {
			return null;
		}
		for (LinkData linkData : links) {
			if (linkData.getUrl() != null) {
				return linkData.getUrl();
			}
		}
		return null;
	}
}
