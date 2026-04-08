package chineseTextService

import java.math.RoundingMode

class ChineseNumberConverter {
    static final String LING_NUMBER = "零"
    static final String DECIMAL_SEPARATOR = "点"
    static final String TEN_THOUSAND_FACTOR = "万"
    static final String TEN_FACTOR = "十"
    static final Map PURE_ARABIC_BY_CHINESE_NUMBER = [
            // Ling for decimal only
            "零" : 0,
            "一" : 1,
            "二" : 2,
            "三" : 3,
            "四" : 4,
            "五" : 5,
            "六" : 6,
            "七" : 7,
            "八" : 8,
            "九" : 9
    ]

    static final Map FACTOR_BY_CHINESE_NUMBER = [
            "十" : 10,
            "百" : 100,
            "千" : 1000,
            "万" : 10000
    ]

    // 一百零一万二千一百一十 -- 1,012,110
    static BigDecimal convertToArabicNumber(String chineseNumber) {
        LinkedList components = chineseNumber?.split(DECIMAL_SEPARATOR)
        String integerPart = components.pollFirst()

        BigDecimal integerNumber = convertIntegerPartToArabicNumber(integerPart)
        if (!components) {
            return integerNumber.setScale(0, RoundingMode.HALF_UP)
        }

        String decimalPart = components.pollFirst()
        BigDecimal decimalNumber = convertDecimalPartToArabicNumber(decimalPart)

        return decimalNumber.add(integerNumber)
    }

    private static BigDecimal convertIntegerPartToArabicNumber(String integerNumber) {
        BigDecimal aNum = 0
        String last = null
        int tenThousandCount = 0
        BigDecimal currentTenThousandPower = 10000.0.pow(tenThousandCount)

        for (int i = integerNumber.length() - 1; i >= 0; i--) {
            String cur = integerNumber[i].toString()
            if (cur == LING_NUMBER) {
                continue
            }
            BigDecimal currentAsNumber = PURE_ARABIC_BY_CHINESE_NUMBER[cur]
            BigDecimal lastAsFactor = FACTOR_BY_CHINESE_NUMBER[last]
            BigDecimal currentAsFactor = FACTOR_BY_CHINESE_NUMBER[cur]

            if (currentAsNumber) {
                if (lastAsFactor) {
                    aNum += currentAsNumber * lastAsFactor * currentTenThousandPower
                    if (last == TEN_THOUSAND_FACTOR) {
                        currentTenThousandPower = 10000.0.pow(++ tenThousandCount)
                    }
                }
                else {
                    aNum += currentAsNumber * currentTenThousandPower
                }
            }
            else if (currentAsFactor && i == 0 && cur == TEN_FACTOR) {
                aNum += currentAsFactor
            }
            last = cur
        }

        return aNum
    }

    private static BigDecimal convertDecimalPartToArabicNumber(String decimalPart) {
        BigDecimal decimal = 0
        int decimalCount = 0
        for (int i = 0; i < decimalPart.length(); i++) {
            String cur = decimalPart[i]
            BigDecimal decimalComponent = PURE_ARABIC_BY_CHINESE_NUMBER[cur]
            decimal += decimalComponent / 10.0.pow(++ decimalCount)
        }

        return decimal.setScale(decimalCount, RoundingMode.HALF_UP)
    }
}
