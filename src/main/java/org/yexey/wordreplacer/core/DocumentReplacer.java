package org.yexey.wordreplacer.core;

import java.util.Map;

public interface DocumentReplacer {
    /**
     * Replace a single bookmark with its replacement
     */
    void replace(String bookmark, String replacement);
    
    /**
     * Replace multiple bookmarks with their replacements
     */
    void replace(Map<String, String> replacements);
}