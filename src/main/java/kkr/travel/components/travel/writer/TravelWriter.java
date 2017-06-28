package kkr.travel.components.travel.writer;

import java.util.Collection;

import kkr.common.errors.BaseException;
import kkr.travel.components.travel.data.TravelData;

public interface TravelWriter {
	void writeTravelData(Collection<TravelData> travelData) throws BaseException;
}
