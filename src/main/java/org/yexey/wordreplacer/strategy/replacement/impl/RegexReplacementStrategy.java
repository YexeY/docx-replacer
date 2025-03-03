package org.yexey.wordreplacer.strategy.replacement.impl;

import org.yexey.wordreplacer.strategy.replacement.ReplacementStrategy;

/**
 * Replacement strategy that handles regex patterns
 */
public class RegexReplacementStrategy implements ReplacementStrategy {
    @Override
    public String replace(String text, String bookmark, String replacement) {
        return text.replaceAll(bookmark, replacement != null ? replacement : "");
    }
}