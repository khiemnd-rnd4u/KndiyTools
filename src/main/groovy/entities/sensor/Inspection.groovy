package entities.sensor


import java.time.ZonedDateTime

class Inspection {
    private TreeMap<String, LogTagSensor> logTagSensorByLocation
    private LogTagSensor environmentSensor
    private TreeSet<ZonedDateTime> dateTimeStamps

    private String id
    private String name
    private String address

    private BigDecimal minTemperatureTolerance
    private BigDecimal maxTemperatureTolerance

    private BigDecimal minHumidityTolerance
    private BigDecimal maxHumidityTolerance

    private BigDecimal allTimeMinTemperature
    private BigDecimal allTimeMaxTemperature

    private BigDecimal allTimeMinHumidity
    private BigDecimal allTimeMaxHumidity

    private Integer sensorCount
    private Integer maxDataCount

    private ZonedDateTime fromTime
    private ZonedDateTime toTime

    static String NONE_CONFORMED_MESSAGE = "KHÔNG ĐẠT/\nNone Conformed"
    static String CONFORMED_MESSAGE = "ĐẠT/\nConformed"

    Inspection(String id,
               String name,
               String address,
               BigDecimal minTemperatureTolerance,
               BigDecimal maxTemperatureTolerance,
               BigDecimal minHumidityTolerance,
               BigDecimal maxHumidityTolerance) {
        this.id = id
        this.name = name
        this.address = address
        this.minTemperatureTolerance = minTemperatureTolerance
        this.maxTemperatureTolerance = maxTemperatureTolerance

        this.minHumidityTolerance = minHumidityTolerance
        this.maxHumidityTolerance = maxHumidityTolerance

        this.fromTime = null
        this.toTime = null

        logTagSensorByLocation = new TreeMap<>()

        sensorCount = 0
        maxDataCount = 0
    }

    LogTagSensor addLogTagSensor(LogTagSensor sensor) {
        String location = sensor.getLocation()
        String altitudeLevel = sensor.getAltitudeLevel()

        logTagSensorByLocation[ getLocationKey(location, altitudeLevel) ] = sensor
        resolveIndicators(sensor)

        return getLogTagSensor(location, altitudeLevel)
    }

    LogTagSensor getLogTagSensor(String location,
                                 String altitudeLevel = null) {
        return logTagSensorByLocation[ getLocationKey(location, altitudeLevel) ]
    }

    LogTagSensor addEnvironmentSensor(LogTagSensor envSensor) {
        environmentSensor = envSensor
        sensorCount = sensorCount + 1

        return environmentSensor
    }

    void resolveIndicators(LogTagSensor logTagSensor) {
        allTimeMinTemperature = LogTagSensor.getMin(logTagSensor.getMinTemperature(), allTimeMinTemperature)
        allTimeMaxTemperature = LogTagSensor.getMax(logTagSensor.getMaxTemperature(), allTimeMaxTemperature)

        allTimeMinHumidity = LogTagSensor.getMin(logTagSensor.getMinHumidity(), allTimeMinHumidity)
        allTimeMaxHumidity = LogTagSensor.getMax(logTagSensor.getMaxHumidity(), allTimeMaxHumidity)

        TreeSet<ZonedDateTime> dateTimeStamps = logTagSensor.getLogTagReadingByDateTime()?.keySet()
        ZonedDateTime fromTime = dateTimeStamps.getFirst()
        ZonedDateTime toTime = dateTimeStamps.getLast()

        this.fromTime = this.fromTime ? [ this.fromTime, fromTime ].min() : fromTime
        this.toTime = this.toTime ? [ this.toTime, fromTime ].max() : toTime

        Integer currentSensorDataSize = logTagSensor.getLogTagReadingByDateTime()?.size()
        maxDataCount = maxDataCount ? [ maxDataCount, currentSensorDataSize ].max() : currentSensorDataSize

        sensorCount = sensorCount + 1
    }

    static String getLocationKey(String location,
                                 String altitudeLevel = null) {
        String key = location
        if (altitudeLevel) {
            key += "-${altitudeLevel}"
        }

        return key
    }

    static String getConformedState(Inspection inspection, LogTagSensor logTagSensor, boolean isTemp) {
        BigDecimal max = isTemp ? logTagSensor.getMaxTemperature() : logTagSensor.getMaxHumidity()
        BigDecimal min = isTemp ? logTagSensor.getMinTemperature() : logTagSensor.getMinHumidity()

        LogTagSensor envTag = inspection.getEnvironmentSensor()
        if (logTagSensor == envTag) {
            return ""
        }

        BigDecimal maxTolerance = isTemp ? inspection.getMaxTemperatureTolerance() : inspection.getMaxHumidityTolerance()
        BigDecimal minTolerance = isTemp ? inspection.getMinTemperatureTolerance() : inspection.getMinHumidityTolerance()

        if (max > maxTolerance) {
            return NONE_CONFORMED_MESSAGE
        }
        if (min < minTolerance) {
            return NONE_CONFORMED_MESSAGE
        }

        return CONFORMED_MESSAGE
    }

    TreeMap<String, LogTagSensor> getLogTagSensorByLocation() {
        return logTagSensorByLocation
    }

    LogTagSensor getEnvironmentSensor() {
        return environmentSensor
    }

    List<ZonedDateTime> getDateTimeStamps() {
        return dateTimeStamps
    }

    String getId() {
        return id
    }

    String getName() {
        return name
    }

    String getAddress() {
        return address
    }

    BigDecimal getMinTemperatureTolerance() {
        return minTemperatureTolerance
    }

    void setMinTemperatureTolerance(BigDecimal minTemperatureTolerance) {
        this.minTemperatureTolerance = minTemperatureTolerance
    }

    BigDecimal getMaxTemperatureTolerance() {
        return maxTemperatureTolerance
    }

    void setMaxTemperatureTolerance(BigDecimal maxTemperatureTolerance) {
        this.maxTemperatureTolerance = maxTemperatureTolerance
    }

    BigDecimal getMinHumidityTolerance() {
        return minHumidityTolerance
    }

    void setMinHumidityTolerance(BigDecimal minHumidityTolerance) {
        this.minHumidityTolerance = minHumidityTolerance
    }

    BigDecimal getMaxHumidityTolerance() {
        return maxHumidityTolerance
    }

    void setMaxHumidityTolerance(BigDecimal maxHumidityTolerance) {
        this.maxHumidityTolerance = maxHumidityTolerance
    }

    BigDecimal getAllTimeMinTemperature() {
        return allTimeMinTemperature
    }

    void setAllTimeMinTemperature(BigDecimal allTimeMinTemperature) {
        this.allTimeMinTemperature = allTimeMinTemperature
    }

    BigDecimal getAllTimeMaxTemperature() {
        return allTimeMaxTemperature
    }

    void setAllTimeMaxTemperature(BigDecimal allTimeMaxTemperature) {
        this.allTimeMaxTemperature = allTimeMaxTemperature
    }

    BigDecimal getAllTimeMinHumidity() {
        return allTimeMinHumidity
    }

    void setAllTimeMinHumidity(BigDecimal allTimeMinHumidity) {
        this.allTimeMinHumidity = allTimeMinHumidity
    }

    BigDecimal getAllTimeMaxHumidity() {
        return allTimeMaxHumidity
    }

    void setAllTimeMaxHumidity(BigDecimal allTimeMaxHumidity) {
        this.allTimeMaxHumidity = allTimeMaxHumidity
    }

    ZonedDateTime getFromTime() {
        return fromTime
    }

    ZonedDateTime getToTime() {
        return toTime
    }

    Integer getSensorCount() {
        return sensorCount
    }

    Integer getMaxDataCount() {
        return maxDataCount
    }

    @Override
    String toString() {
        return "Inspection{" +
                "logTagSensorByLocation=" + logTagSensorByLocation +
                ", environmentSensor=" + environmentSensor +
                ", dateTimeStamps=" + dateTimeStamps +
                ", id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", address='" + address + '\'' +
                ", minTemperatureTolerance=" + minTemperatureTolerance +
                ", maxTemperatureTolerance=" + maxTemperatureTolerance +
                ", minHumidityTolerance=" + minHumidityTolerance +
                ", maxHumidityTolerance=" + maxHumidityTolerance +
                ", allTimeMinTemperature=" + allTimeMinTemperature +
                ", allTimeMaxTemperature=" + allTimeMaxTemperature +
                ", allTimeMinHumidity=" + allTimeMinHumidity +
                ", allTimeMaxHumidity=" + allTimeMaxHumidity +
                ", fromTime=" + fromTime +
                ", toTime=" + toTime +
                '}';
    }
}
