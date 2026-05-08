package services.kndiyLibraries

class KndiyFormatter {
    static String getStringWithoutRedundantSpaces(String input) {
        if (!input) {
            return input
        }
        input = input.replace("  ", " ")
        input = input.replace("  ", " ")
        input = input.replace("  ", " ")
        input = input.replace("  ", " ")
        return input.trim().split(" ")
                .collect { String component ->
                    return component.trim()
                }
                .join(" ")
    }

}
