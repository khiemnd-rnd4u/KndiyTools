package services.kndiyLibraries

import java.time.ZoneId
import java.time.ZonedDateTime
import java.time.format.DateTimeFormatter

class DateTimeResolver {
    static final Map MONTH_BY_ABBREVIATION = [
            "jan" : 1,
            "feb" : 2,
            "mar" : 3,
            "apr" : 4,
            "may" : 5,
            "jun" : 6,
            "jul" : 7,
            "aug" : 8,
            "sep" : 9,
            "oct" : 10,
            "nov" : 11,
            "dec" : 12,
    ]
    static final Map DAYS_BY_MONTH = [
            1 : 31,
            2 : 28,
            3 : 31,
            4 : 30,
            5 : 31,
            6 : 30,
            7 : 31,
            8 : 31,
            9 : 30,
            10 : 31,
            11 : 30,
            12 : 31,
    ]
    static final Map DAYS_BY_MONTH_LEAP = [
            1 : 31,
            2 : 29,
            3 : 31,
            4 : 30,
            5 : 31,
            6 : 30,
            7 : 31,
            8 : 31,
            9 : 30,
            10 : 31,
            11 : 30,
            12 : 31,
    ]

    static String DEFAULT_PATTERN = "yyyy-MM-dd'T'HH:mm:ss"
    static String DATE_PATTERN = "yyyy-MM-dd"
    static String TIME_PATTERN = "HH:mm:ss"
    static ZoneId ZONE_ID_HCM = ZoneId.of("Asia/Ho_Chi_Minh")

    static ZonedDateTime getZonedDateTime(String dateTime,
                                          String zone = null) {
        if (dateTime.length() == 10) {
            dateTime += "T00:00:00"
        }

        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(DEFAULT_PATTERN)
                .withZone(getZoneId(zone))

        return ZonedDateTime.parse(dateTime, dateTimeFormatter)
    }

    static String getDateTimeString(ZonedDateTime zonedDateTime,
                                    String pattern = null,
                                    String zone = null) {

        return zonedDateTime?.withZoneSameInstant(getZoneId(zone))
                ?.format(pattern ?: DEFAULT_PATTERN)
    }

    static String getDateString(ZonedDateTime zonedDateTime,
                                    String pattern = null,
                                    String zone = null) {

        return zonedDateTime?.withZoneSameInstant(getZoneId(zone))
                ?.format(pattern ?: DATE_PATTERN)
    }

    static String getTimeString(ZonedDateTime zonedDateTime,
                                    String pattern = null,
                                    String zone = null) {

        return zonedDateTime?.withZoneSameInstant(getZoneId(zone))
                ?.format(pattern ?: TIME_PATTERN)
    }

    static ZoneId getZoneId(String zone) {
        return (zone
                ? ZoneId.of(zone)
                : ZONE_ID_HCM)
    }
}
