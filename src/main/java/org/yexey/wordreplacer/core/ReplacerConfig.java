package org.yexey.wordreplacer.core;

import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;
import org.yexey.wordreplacer.strategy.replacement.ReplacementStrategy;
import org.yexey.wordreplacer.strategy.replacement.impl.StandardReplacementStrategy;

/**
 * Configuration class for the document replacer
 */
@Getter
@Setter
@Accessors(chain = true)
public class ReplacerConfig {
    private ReplacementStrategy strategy = new StandardReplacementStrategy();
    private boolean processHeaders = true;
    private boolean processFooters = true;
}