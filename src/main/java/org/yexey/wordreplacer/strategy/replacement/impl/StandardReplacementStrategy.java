package org.yexey.wordreplacer.strategy.replacement.impl;

import org.yexey.wordreplacer.strategy.replacement.ReplacementStrategy;

public class StandardReplacementStrategy implements ReplacementStrategy {
    @Override
    public String replace(String text, String bookmark, String replacement) {
        return text.replace(bookmark, replacement != null ? replacement : "");
    }
}