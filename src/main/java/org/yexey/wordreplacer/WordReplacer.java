package org.yexey.wordreplacer;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.yexey.wordreplacer.internal.strategy.tracker.ReplacementTracker;
import org.yexey.wordreplacer.internal.strategy.tracker.impl.SimpleReplacementTracker;
import org.yexey.wordreplacer.internal.strategy.visitor.DocumentElementVisitor;
import org.yexey.wordreplacer.internal.strategy.visitor.impl.BookmarkFinderVisitor;
import org.yexey.wordreplacer.internal.strategy.visitor.impl.RemovalVisitor;
import org.yexey.wordreplacer.internal.strategy.visitor.impl.ReplacementVisitor;

import java.util.Map;
import java.util.Optional;

import static org.apache.commons.lang3.StringUtils.isBlank;

/**
 * Main implementation of the document replacer
 */
@Slf4j
public class WordReplacer implements WordReplacerIF {

    private final XWPFDocument document;
    @Getter
    private final ReplacementTracker tracker;

    public WordReplacer(XWPFDocument document) {
        this.document = document;
        this.tracker = new SimpleReplacementTracker();
    }

    @Override
    public void replace(String bookmark, String replacement) {
        // Create a document visitor for this operation
        ReplacementVisitor visitor = new ReplacementVisitor(
                bookmark,
                replacement,
                tracker);

        // Process document elements
        processDocument(visitor);
    }

    @Override
    public void replace(Map<String, String> replacements) {
        // Process each replacement
        for (Map.Entry<String, String> entry : replacements.entrySet()) {
            replace(entry.getKey(), entry.getValue());
        }
    }

    /**
     * Replace with default text if replacement is empty
     */
    @Override
    public void replaceOrDefault(String bookmark, String replacement, String defaultText) {
        replace(bookmark, (replacement == null || replacement.isEmpty()) ? defaultText : replacement);
    }

    /**
     * Process optional replacement
     */
    @Override
    public void replace(String bookmark, Optional<String> replacement) {
        replace(bookmark, replacement.orElse(""));
    }

    /**
     * Remove document paragraphs containing the bookmark
     */
    @Override
    public void removeParagraph(String bookmark) {
        if (isBlank(bookmark)) {
            return;
        }
        // Create a removal visitor
        RemovalVisitor visitor = new RemovalVisitor(document, bookmark);

        // Process document elements
        processDocument(visitor);
    }

    /**
     * Check if a bookmark exists in the document
     */
    @Override
    public boolean hasBookmark(String bookmark) {
        BookmarkFinderVisitor finder = new BookmarkFinderVisitor(bookmark);
        processDocument(finder);
        return finder.isFound();
    }

    /**
     * Process the document with a visitor
     */
    private void processDocument(DocumentElementVisitor visitor) {
        // Process paragraphs in the document body
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            visitor.visitParagraph(paragraph);
        }

        // Process tables in the document
        for (XWPFTable table : document.getTables()) {
            visitor.visitTable(table);
        }

        // Process headers if configured
        for (XWPFHeader header : document.getHeaderList()) {
            visitor.visitHeader(header);
        }

        // Process footers if configured
        for (XWPFFooter footer : document.getFooterList()) {
            visitor.visitFooter(footer);
        }
    }
}