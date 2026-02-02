package novelParsingService

class NovelParser implements ParserInterface {
    String fileAbsolutePath
    String outputFileAbsolutePath

    NovelParser(String fileAbsolutePath) {
        this.fileAbsolutePath = fileAbsolutePath
    }

    @Override
    List<String> getParagraphTexts() {
        return null
    }

    @Override
    void writeNewDocument(String outputFileAbsolutePath, TreeMap paragraphsByTypeByChapterNumber) {
        this.outputFileAbsolutePath = outputFileAbsolutePath
    }
}