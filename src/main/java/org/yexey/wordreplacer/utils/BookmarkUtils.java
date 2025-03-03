package org.yexey.wordreplacer.utils;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.List;

public class BookmarkUtils {
    /**
     * Checks if a paragraph contains the specified bookmark.
     *
     * @param paragraph the paragraph to check
     * @param bookmark  the bookmark to search for
     * @return true if the paragraph contains the bookmark, false otherwise
     */
    public static boolean containsBookmark(XWPFParagraph paragraph, String bookmark) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty()) {
            return false;
        }

        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text != null) {
                sb.append(text);
            }
        }

        String paragraphText = sb.toString();
        return paragraphText.contains(bookmark);
    }
}