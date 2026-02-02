package novelParsingService

import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph

class NovelParserXwpf extends NovelParser {
    NovelParserXwpf(String fileAbsolutePath) {
        super(fileAbsolutePath)
    }

    @Override
    List<String> getParagraphTexts() {
        XWPFDocument document = new XWPFDocument(new FileInputStream(fileAbsolutePath))
        List paragrapths = document?.getParagraphs()?.collect { XWPFParagraph paragraph ->
            return paragraph.getText().trim()
        }
        document.close()
        return paragrapths
    }

}
