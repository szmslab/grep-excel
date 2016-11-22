/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel.cli;

import org.kohsuke.args4j.Argument;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.Option;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * {@link CmdLineParser}によりパースされたコマンドライン引数を保持するクラスです。
 *
 * @author szmslab
 */
class CommandLineParameter {

    /**
     * {@link Argument#metaVar()}の設定値（検索パターン）。
     */
    static final String META_VAR_PATTERN_TEXT = "PATTERN";

    /**
     * {@link Argument#metaVar()}の設定値（ファイル・ディレクトリパスのリスト）。
     */
    static final String META_VAR_PATH_LIST = "FILE";

    /**
     * 引数（検索パターン）。
     */
    @Argument(metaVar = META_VAR_PATTERN_TEXT, required = true, hidden = true)
    String patternText;

    /**
     * 引数（ファイル・ディレクトリパスのリスト）。
     */
    @Argument(index = 1, metaVar = META_VAR_PATH_LIST, required = true, handler = ExistingPathOptionHandler.class, hidden = true)
    List<Path> pathList = new ArrayList<>();

    /**
     * オプション（数式セルの計算結果を検索する）。
     */
    @Option(name = "-f", aliases = {"--formula-result"}, usage = "search for calculated result of formula")
    boolean formulaResult;

    /**
     * オプション（ヘルプを表示する）。
     */
    @Option(name = "-h", aliases = {"--help"}, usage = "display help information and exit", help = true)
    boolean help;

    /**
     * オプション（大文字・小文字を区別しない）。
     */
    @Option(name = "-i", aliases = {"--ignore-case"}, usage = "ignore case distinctions")
    boolean ignoreCase;

    /**
     * オプション（リテラル構文解析を有効にする）。
     */
    @Option(name = "-l", aliases = {"--literal"}, usage = "enable literal parsing of the pattern")
    boolean literal;

    /**
     * オプション（検索処理を並列実行する）。
     */
    @Option(name = "-p", aliases = {"--parallel"}, usage = "perform search processing in parallel")
    boolean parallel;

    /**
     * オプション（ディレクトリを再帰的に検索する）。
     */
    @Option(name = "-r", aliases = {"--recursive"}, usage = "search directories for file recursively")
    boolean recursive;

    /**
     * オプション（集計を出力する）。
     */
    @Option(name = "-s", aliases = {"--summary"}, usage = "print a summary of results")
    boolean summary;

    /**
     * オプション（バージョンを表示する）。
     */
    @Option(name = "-v", aliases = {"--version"}, usage = "display version information and exit", help = true)
    boolean version;

    @Override
    public String toString() {
        return "{" +
                "patternText=" + patternText +
                ", pathList=" + pathList +
                ", formulaResult=" + formulaResult +
                ", help=" + help +
                ", ignoreCase=" + ignoreCase +
                ", literal=" + literal +
                ", parallel=" + parallel +
                ", recursive=" + recursive +
                ", summary=" + summary +
                ", version=" + version +
                "}";
    }

}
