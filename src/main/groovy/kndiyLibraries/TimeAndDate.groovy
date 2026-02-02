package kndiyLibraries

class TimeAndDate {
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
}
