package cache

class Cache {
    static Set errorMessages

    static void initiateCache() {
        errorMessages = [ ]
    }

    static void printErrorMessages() {
        List messages = (errorMessages as List).sort()
        for (String message : messages) {
            println(message)
        }
    }
}