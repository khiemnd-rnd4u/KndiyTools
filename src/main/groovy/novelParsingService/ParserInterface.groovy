package novelParsingService

interface ParserInterface {
    List<String> getParagraphTexts()

    void writeNewDocument(String outputFileAbsolutePath, TreeMap paragraphsByTypeByChapterNumber)
}
