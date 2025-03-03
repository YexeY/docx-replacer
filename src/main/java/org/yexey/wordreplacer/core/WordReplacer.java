package org.yexey.wordreplacer.core;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.yexey.wordreplacer.strategy.replacement.ReplacementStrategy;
import org.yexey.wordreplacer.strategy.tracker.ReplacementTracker;
import org.yexey.wordreplacer.strategy.tracker.impl.SimpleReplacementTracker;
import org.yexey.wordreplacer.strategy.visitor.DocumentElementVisitor;
import org.yexey.wordreplacer.strategy.visitor.impl.BookmarkFinderVisitor;
import org.yexey.wordreplacer.strategy.visitor.impl.RemovalVisitor;
import org.yexey.wordreplacer.strategy.visitor.impl.ReplacementVisitor;

import java.util.Map;
import java.util.Optional;

import static org.apache.commons.lang3.StringUtils.isBlank;


/**
 * Main implementation of the document replacer
 */
@Slf4j
public class WordReplacer implements DocumentReplacer {

    private final XWPFDocument document;
    private final ReplacerConfig config;
    private final ReplacementTracker tracker;

    /**
     * Main constructor with default configuration
     */
    public WordReplacer(XWPFDocument document) {
        this(document, new ReplacerConfig(), new SimpleReplacementTracker());
    }

    /**
     * Constructor with custom configuration
     */
    public WordReplacer(XWPFDocument document, ReplacerConfig config, ReplacementTracker tracker) {
        this.document = document;
        this.config = config;
        this.tracker = tracker;
    }

    @Override
    public void replace(String bookmark, String replacement) {
        // Create a document visitor for this operation
        ReplacementVisitor visitor = new ReplacementVisitor(
                bookmark,
                replacement,
                config.getStrategy(),
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
    public void replaceOrDefault(String bookmark, String replacement, String defaultText) {
        replace(bookmark, (replacement == null || replacement.isEmpty()) ? defaultText : replacement);
    }

    public void removeParagraph(String bookmark) {
        if (!isBlank(bookmark)) {
            removeElementsContaining(bookmark);
        }
    }

    /**
     * Replace or remove paragraph if replacement is empty
     */
    public void replaceOrRemoveParagraph(String bookmark, String replacement) {
        if (isBlank(replacement)) {
            removeElementsContaining(bookmark);
        } else {
            replace(bookmark, replacement);
        }
    }

    /**
     * Process optional replacement
     */
    public void replace(String bookmark, Optional<String> replacement) {
        replace(bookmark, replacement.orElse(""));
    }

    /**
     * Remove document elements containing the bookmark
     */
    public void removeElementsContaining(String bookmark) {
        // Create a removal visitor
        RemovalVisitor visitor = new RemovalVisitor(document, bookmark);

        // Process document elements
        processDocument(visitor);
    }

    /**
     * Get statistics about replacements
     */
    public ReplacementTracker getTracker() {
        return tracker;
    }

    /**
     * Check if a bookmark exists in the document
     */
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
        if (config.isProcessHeaders()) {
            for (XWPFHeader header : document.getHeaderList()) {
                visitor.visitHeader(header);
            }
        }

        // Process footers if configured
        if (config.isProcessFooters()) {
            for (XWPFFooter footer : document.getFooterList()) {
                visitor.visitFooter(footer);
            }
        }
    }

    /**
     * Builder class for fluent creation of WordReplacer
     */
    public static class Builder {
        private final XWPFDocument document;
        private final ReplacerConfig config = new ReplacerConfig();
        private ReplacementTracker tracker = new SimpleReplacementTracker();

        public Builder(XWPFDocument document) {
            this.document = document;
        }

        public Builder withStrategy(ReplacementStrategy strategy) {
            config.setStrategy(strategy);
            return this;
        }

        public Builder processHeaders(boolean process) {
            config.setProcessHeaders(process);
            return this;
        }

        public Builder processFooters(boolean process) {
            config.setProcessFooters(process);
            return this;
        }

        public Builder withTracker(ReplacementTracker tracker) {
            this.tracker = tracker;
            return this;
        }

        public WordReplacer build() {
            return new WordReplacer(document, config, tracker);
        }
    }
}