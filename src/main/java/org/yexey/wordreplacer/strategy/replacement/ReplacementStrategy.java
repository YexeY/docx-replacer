package org.yexey.wordreplacer.strategy.replacement;

/**
 * Strategy interface for different replacement operations
 */
public interface ReplacementStrategy {
    /**
     * Performs the replacement in the given text
     */
    String replace(String text, String bookmark, String replacement);
}