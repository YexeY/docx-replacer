package io.github.yexey.wordreplacer.internal.strategy.visitor.impl;

import org.apache.poi.xwpf.usermodel.*;
import io.github.yexey.wordreplacer.internal.strategy.tracker.ReplacementTracker;
import io.github.yexey.wordreplacer.internal.strategy.visitor.DocumentElementVisitor;

import java.util.List;

/**
 * Visitor for replacing bookmarks
 */
public class ReplacementVisitor implements DocumentElementVisitor {
    private final String bookmark;
    private final String replacement;
    private final ReplacementTracker tracker;

    public ReplacementVisitor(String bookmark, String replacement, ReplacementTracker tracker) {
        this.bookmark = bookmark;
        this.replacement = replacement;
        this.tracker = tracker;
    }

    @Override
    public void visitParagraph(XWPFParagraph paragraph) {
        boolean success = replaceInParagraph(paragraph, bookmark, replacement);
        tracker.trackReplacement(bookmark, replacement, success);
    }

    @Override
    public void visitTable(XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            visitTableRow(row);
        }
    }

    @Override
    public void visitTableCell(XWPFTableCell cell) {
        // Process paragraphs within the cell
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            visitParagraph(paragraph);
        }

        // Process nested tables
        for (XWPFTable nestedTable : cell.getTables()) {
            visitTable(nestedTable);
        }
    }

    @Override
    public void visitTableRow(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            visitTableCell(cell);
        }
    }

    @Override
    public void visitHeader(XWPFHeader header) {
        for (XWPFParagraph paragraph : header.getParagraphs()) {
            visitParagraph(paragraph);
        }

        for (XWPFTable table : header.getTables()) {
            visitTable(table);
        }
    }

    @Override
    public void visitFooter(XWPFFooter footer) {
        for (XWPFParagraph paragraph : footer.getParagraphs()) {
            visitParagraph(paragraph);
        }

        for (XWPFTable table : footer.getTables()) {
            visitTable(table);
        }
    }

    /**
     * Handles the case where a bookmark spans multiple runs in a paragraph.
     *
     * @param paragraph   the paragraph containing the runs
     * @param bookmark    the bookmark to search for
     * @param replacement the text to replace the bookmark with
     * @return true if the bookmark was found and replaced, false otherwise
     */
    private boolean replaceInParagraph(XWPFParagraph paragraph, String bookmark, String replacement) {
        // Implementation similar to original code but using the strategy
        // This is a simplified version - would need formatting preservation logic

        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) {
            return false;
        }

        StringBuilder fullText = new StringBuilder();
        int[] runEndPositions = new int[runs.size() + 1];
        runEndPositions[0] = 0;

        // Build full text and track positions
        for (int i = 0; i < runs.size(); i++) {
            String text = runs.get(i).getText(0);
            if (text == null) text = "";
            fullText.append(text);
            runEndPositions[i + 1] = fullText.length();
        }

        // Find bookmark in full text
        String completeText = fullText.toString();
        int bookmarkStart = completeText.indexOf(bookmark);
        if (bookmarkStart == -1) {
            return false;
        }

        int bookmarkEnd = bookmarkStart + bookmark.length();

        // Find which runs contain the bookmark
        int startRunIndex = -1;
        int endRunIndex = -1;

        for (int i = 0; i < runs.size(); i++) {
            if (startRunIndex == -1 && runEndPositions[i + 1] > bookmarkStart) {
                startRunIndex = i;
            }
            if (endRunIndex == -1 && runEndPositions[i + 1] >= bookmarkEnd) {
                endRunIndex = i;
                break;
            }
        }

        if (startRunIndex == -1 || endRunIndex == -1) {
            return false;
        }

        // Modify the runs
        XWPFRun startRun = runs.get(startRunIndex);
        String startRunText = startRun.getText(0) != null ? startRun.getText(0) : "";
        int bookmarkStartInRun = bookmarkStart - runEndPositions[startRunIndex];

        XWPFRun endRun = runs.get(endRunIndex);
        String endRunText = endRun.getText(0) != null ? endRun.getText(0) : "";
        int bookmarkEndInRun = bookmarkEnd - runEndPositions[endRunIndex];

        // Create new text for first run
        String textBeforeBookmark = startRunText.substring(0, bookmarkStartInRun);
        String textAfterBookmark = endRunText.substring(bookmarkEndInRun);
        String newStartRunText = textBeforeBookmark + replacement + textAfterBookmark;

        // Set the new text and remove extra runs
        startRun.setText(newStartRunText, 0);
        for (int i = endRunIndex; i > startRunIndex; i--) {
            paragraph.removeRun(i);
        }

        return true;
    }
}