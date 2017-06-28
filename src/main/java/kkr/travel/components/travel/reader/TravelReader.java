package kkr.travel.components.travel.reader;

import java.util.Collection;

import kkr.common.errors.BaseException;
import kkr.travel.components.travel.data.TravelData;

public interface TravelReader {

	Collection<TravelData> readTravelData() throws BaseException;
}
