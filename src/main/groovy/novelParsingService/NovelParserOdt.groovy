package novelParsingService

import org.odftoolkit.odfdom.doc.OdfDocument

class NovelParserOdt extends NovelParser {
    NovelParserOdt(String fileAbsolutePath) {
        super(fileAbsolutePath)
    }

    @Override
    List<String> getParagraphTexts() {
        try (OdfDocument document = OdfDocument.loadDocument(fileAbsolutePath)) {
            return document.getContentRoot().collect {
                String paragraph = it.getTextContent()
                return paragraph
            }
        }
        catch (ignored) {
            return null
        }
    }
}
