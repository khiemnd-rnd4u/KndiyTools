package entities.sensor

import services.kndiyLibraries.DateTimeResolver

import java.time.ZoneId
import java.time.ZonedDateTime

class LogTagSensor {
    private static final Map SENSOR_SERIAL_BY_ID = [
            "51" : "1010209722",
            "52" : "1010209607",
            "53" : "1010211200",
            "54" : "1010211214",
            "56" : "1010209701",
            "57" : "1010211181",
            "59" : "1010209606",
            "60" : "1010211221",
            "61" : "1010209712",
            "62" : "1010211201",
            "63" : "1010209681",
            "64" : "1010211198",
            "65" : "1010211186",
            "66" : "1010209605",
            "67" : "1010211218",
            "68" : "1010209676",
            "69" : "1010211183",
            "70" : "1010211199",
            "71" : "1010211210",
            "72" : "1010211207",
            "73" : "1010211209",
            "74" : "1010211179",
            "75" : "1010211180",
            "76" : "1010209708",
            "77" : "1010211182",
            "78" : "1010209710",
            "79" : "1010211196",
    ]
    private static Map SENSOR_ID_BY_SERIAL = SENSOR_SERIAL_BY_ID
            .collectEntries { String number, String serial ->
                return [
                        (serial) : number
                ]
            }

    private final String id
    private final String serial
    private final String location
    private final String altitudeLevel
    private final LogTagCalibration calibration
    private final TreeMap<ZonedDateTime, LogTagReading> logTagReadingByDateTime

    private BigDecimal minTemperature
    private ZonedDateTime minTemperatureTime

    private BigDecimal minHumidity
    private ZonedDateTime minHumidityTime

    private BigDecimal maxTemperature
    private ZonedDateTime maxTemperatureTime

    private BigDecimal maxHumidity
    private ZonedDateTime maxHumidityTime

    private BigDecimal sumTemperature
    private BigDecimal sumHumidity

    private BigDecimal avgTemperature
    private BigDecimal avgHumidity

    LogTagSensor(String idOrSerial,
                 String location,
                 String altitudeLevel = null,
                 ZonedDateTime inspectionDate = null) {
        if (inspectionDate == null) {
            inspectionDate = ZonedDateTime.now(ZoneId.of("Asia/Ho_Chi_Minh"))
        }

        this.id = SENSOR_ID_BY_SERIAL?.getAt(idOrSerial) ?: idOrSerial
        this.serial = SENSOR_SERIAL_BY_ID?.getAt(idOrSerial) ?: idOrSerial
        this.location = location
        this.altitudeLevel = altitudeLevel

        this.calibration = new LogTagCalibration(serial, inspectionDate)
        this.logTagReadingByDateTime = new TreeMap<>((t1, t2) -> {
            return t1 <=> t2
        })
    }

    // DateTime with format yyyy-MM-dd'T'HH:mm:ss
    void addReading(ZonedDateTime dateTime,
                    BigDecimal temperature,
                    BigDecimal humidity,
                    String event = null) {
        LogTagReading logTagReading = new LogTagReading(dateTime, temperature, humidity, event)
        calculateReferenceIndicators(logTagReading, dateTime)

        logTagReadingByDateTime[ dateTime ] = logTagReading
    }

    LogTagReading getReading(ZonedDateTime dateTime) {
        return logTagReadingByDateTime[ dateTime ]
    }

    void calculateReferenceIndicators(LogTagReading logTagReading, ZonedDateTime dateTime) {
        BigDecimal temperature = logTagReading.getTemperature()
        BigDecimal humidity = logTagReading.getHumidity()

        minTemperature = getMin(temperature, minTemperature)
        minHumidity = getMin(humidity, minHumidity)

        maxTemperature = getMax(temperature, maxTemperature)
        maxHumidity = getMax(humidity, maxHumidity)

        updateMinMaxTime(dateTime, temperature)
        updateMinMaxTime(dateTime, humidity, false)

        sumTemperature = getSum(temperature, sumTemperature)
        sumHumidity = getSum(humidity, sumHumidity)
    }

    void resolveAvg() {
        Integer dataCount = logTagReadingByDateTime?.size()
        avgTemperature = sumTemperature / dataCount
        avgHumidity = sumHumidity / dataCount
    }

    private void updateMinMaxTime(ZonedDateTime dateTime,
                                  BigDecimal value,
                                  boolean isTemperature = true) {
        BigDecimal min = isTemperature ? minTemperature : minHumidity
        BigDecimal max = isTemperature ? maxTemperature : maxHumidity

        if (isTemperature) {
            minTemperatureTime = min == value ? dateTime : minTemperatureTime
            maxTemperatureTime = max == value ? dateTime : maxTemperatureTime
        }
        else {
            minHumidityTime = min == value ? dateTime : minHumidityTime
            maxHumidityTime = max == value ? dateTime : maxHumidityTime
        }
    }

    static BigDecimal getMin(BigDecimal value, BigDecimal min) {
        if (!min) {
            return value
        }
        return [ value, min ].min()
    }

    static BigDecimal getMax(BigDecimal value, BigDecimal max) {
        if (!max) {
            return value
        }
        return [ value, max ].max()
    }

    static BigDecimal getSum(BigDecimal value, BigDecimal sum) {
        if (!sum) {
            return value
        }

        return sum + value
    }

    TreeMap<ZonedDateTime, LogTagReading> getLogTagReadingByDateTime() {
        return logTagReadingByDateTime
    }

    LogTagCalibration getCalibration() {
        return calibration
    }

    String getAltitudeLevel() {
        return altitudeLevel
    }

    String getLocation() {
        return location
    }

    String getSerial() {
        return serial
    }

    String getId() {
        return id
    }

    BigDecimal getMinTemperature() {
        return minTemperature
    }

    ZonedDateTime getMinTemperatureTime() {
        return minTemperatureTime
    }

    BigDecimal getMinHumidity() {
        return minHumidity
    }

    ZonedDateTime getMinHumidityTime() {
        return minHumidityTime
    }

    BigDecimal getMaxTemperature() {
        return maxTemperature
    }

    ZonedDateTime getMaxTemperatureTime() {
        return maxTemperatureTime
    }

    BigDecimal getMaxHumidity() {
        return maxHumidity
    }

    ZonedDateTime getMaxHumidityTime() {
        return maxHumidityTime
    }

    BigDecimal getAvgTemperature() {
        return avgTemperature
    }

    BigDecimal getAvgHumidity() {
        return avgHumidity
    }

    @Override
    String toString() {
        return "LogTagSensor{" +
                "id='" + id + '\'' +
                ", serial='" + serial + '\'' +
                ", location='" + location + '\'' +
                ", altitudeLevel='" + altitudeLevel + '\'' +
                ", calibration=" + calibration +
                ", logTagReadingByDateTime=" + logTagReadingByDateTime +
                ", minTemperature=" + minTemperature +
                ", minTemperatureTime=" + minTemperatureTime +
                ", minHumidity=" + minHumidity +
                ", minHumidityTime=" + minHumidityTime +
                ", maxTemperature=" + maxTemperature +
                ", maxTemperatureTime=" + maxTemperatureTime +
                ", maxHumidity=" + maxHumidity +
                ", maxHumidityTime=" + maxHumidityTime +
                ", avgTemperature=" + sumTemperature +
                ", avgHumidity=" + sumHumidity +
                '}'
    }
}
