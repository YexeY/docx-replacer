package io.github.yexey.wordreplacer;

import io.github.yexey.wordreplacer.internal.strategy.tracker.ReplacementTracker;
import io.github.yexey.wordreplacer.internal.strategy.tracker.impl.SimpleReplacementTracker;
import io.github.yexey.wordreplacer.internal.strategy.visitor.DocumentElementVisitor;
import io.github.yexey.wordreplacer.internal.strategy.visitor.impl.BookmarkFinderVisitor;
import io.github.yexey.wordreplacer.internal.strategy.visitor.impl.RemovalVisitor;
import io.github.yexey.wordreplacer.internal.strategy.visitor.impl.ReplacementVisitor;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.util.Map;
import java.util.Optional;

/**
 * WordReplacer - Main implementation for replacing placeholders in MS Word documents.
 *
 * This class provides functionality to replace placeholders within
 * Word documents with actual content. It uses the visitor pattern to traverse all document elements
 * including paragraphs, tables, headers, and footers.
 *
 * The class supports:
 * - Single replacements
 * - Batch replacements
 * - Default value replacements
 * - Optional value replacements
 * - Paragraph removal based on placeholder content
 * - Tracking of successful and failed replacements
 *
 * Usage example:
 * <pre>
 *     WordReplacer replacer = new WordReplacer(document);
 *     replacer.replace("{{NAME}}", "John Doe");
 *
 *     Map<String, String> replacements = new HashMap<>();
 *     replacements.put("{{EMAIL}}", "john.doe@example.com");
 *     replacements.put("{{PHONE}}", "(555) 123-4567");
 *     replacer.replace(replacements);
 *
 *     replacer.removeParagraph("{{NOTES}}");
 * </pre>
 */
@Slf4j
public class WordReplacer implements WordReplacerIF {

    /**
     * The Word document being processed
     */
    private final XWPFDocument document;

    /**
     * Tracks statistics about replacements performed (success/failure)
     */
    @Getter
    private final ReplacementTracker tracker;

    /**
     * Creates a new WordReplacer for the given document
     *
     * @param document The XWPFDocument to process (MS Word document)
     */
    public WordReplacer(XWPFDocument document) {
        this.document = document;
        this.tracker = new SimpleReplacementTracker();
    }

    /**
     * Replaces a single placeholder with the specified replacement text throughout the document.
     * The replacement is performed in all document elements (paragraphs, tables, headers, footers).
     *
     * @param bookmark The placeholder text to find (typically in {{PLACEHOLDER}} format)
     * @param replacement The text to replace the placeholder with
     */
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

    /**
     * Performs multiple replacements in a single pass through the document.
     * This is more efficient than calling replace() multiple times as the document
     * is only traversed once per entry in the map.
     *
     * @param replacements A map of placeholders to their replacement values
     */
    @Override
    public void replace(Map<String, String> replacements) {
        // Process each replacement
        for (Map.Entry<String, String> entry : replacements.entrySet()) {
            replace(entry.getKey(), entry.getValue());
        }
    }

    /**
     * Replaces a placeholder with the given replacement text, or with a default text
     * if the replacement is null.
     *
     * This is useful when you want to ensure a placeholder is always replaced, even
     * when no specific replacement value is available.
     *
     * @param bookmark The placeholder text to find
     * @param replacement The primary replacement text (can be null)
     * @param defaultText The fallback text to use if replacement is null or empty
     */
    @Override
    public void replaceOrDefault(String bookmark, String replacement, String defaultText) {
        replace(bookmark, replacement == null ? defaultText : replacement);
    }

    /**
     * Replaces a placeholder with an Optional value.
     * If the Optional is empty, the placeholder is replaced with an empty string.
     *
     * @param bookmark The placeholder text to find
     * @param replacement An Optional containing the replacement text
     */
    @Override
    public void replace(String bookmark, Optional<String> replacement) {
        replace(bookmark, replacement.orElse(""));
    }

    /**
     * Completely removes paragraphs containing the specified placeholder.
     *
     * This is useful for conditional sections in templates where certain content
     * should be removed entirely rather than just having the placeholder replaced.
     *
     * @param bookmark The placeholder text to search for
     */
    @Override
    public void removeParagraph(String bookmark) {
        if (StringUtils.isBlank(bookmark)) {
            return;
        }
        // Create a removal visitor
        RemovalVisitor visitor = new RemovalVisitor(document, bookmark);

        // Process document elements
        processDocument(visitor);
    }

    /**
     * Checks if a specific placeholder exists anywhere in the document.
     *
     * @param bookmark The placeholder text to search for
     * @return true if the placeholder exists in the document, false otherwise
     */
    @Override
    public boolean hasBookmark(String bookmark) {
        BookmarkFinderVisitor finder = new BookmarkFinderVisitor(bookmark);
        processDocument(finder);
        return finder.isFound();
    }

    /**
     * Processes the entire document with the specified visitor.
     *
     * This method implements the Visitor pattern to traverse all document elements:
     * - Main document paragraphs
     * - Tables (and their nested paragraphs)
     * - Headers
     * - Footers
     *
     * The visitor is responsible for the actual processing of each element type.
     *
     * @param visitor The DocumentElementVisitor to apply to each element
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