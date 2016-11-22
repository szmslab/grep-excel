/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel.cli;

import com.szmslab.grepexcel.GrepExcel;
import com.szmslab.grepexcel.GrepExcelResultSummary;
import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.ParserProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * {@link GrepExcel}をコマンドラインユーティリティとして実行するクラスです。
 *
 * @author szmslab
 */
class CommandLineRunner {

    /**
     * ロガー。
     */
    private static final Logger LOG = LoggerFactory.getLogger(CommandLineRunner.class);

    /**
     * コマンド名。
     */
    private static final String COMMAND = "grepexcel";

    /**
     * コマンドを実行します。
     *
     * @param args コマンドライン引数
     * @return 正常に処理が終了した場合は {@code 0}
     */
    int run(String... args) {
        LOG.debug("args: {}", Arrays.asList(args));

        CommandLineParameter parameter = new CommandLineParameter();
        CmdLineParser parser = new CmdLineParser(
                parameter,
                ParserProperties.defaults()
                        .withShowDefaults(false)
                        .withUsageWidth(120));

        try {
            parser.parseArgument(args);
        } catch (CmdLineException e) {
            System.err.println(e.getMessage());
            System.err.println();
            help(parser, System.err);
            return 1;
        }

        LOG.debug("parameter: {}", parameter);

        if (parameter.help) {
            help(parser, System.out);
        } else if (parameter.version) {
            version();
        } else {
            grep(parameter);
        }
        return 0;
    }

    /**
     * {@link GrepExcel}を使用して、Excelファイル内の文字列を検索します。
     *
     * @param parameter {@link CmdLineParser}によりパースされたコマンドライン引数
     */
    private void grep(CommandLineParameter parameter) {
        long startTime = System.currentTimeMillis();
        GrepExcelResultSummary summary =
                new GrepExcel()
                        .ignoreCase(parameter.ignoreCase)
                        .literal(parameter.literal)
                        .formulaResult(parameter.formulaResult)
                        .recursive(parameter.recursive)
                        .parallel(parameter.parallel)
                        .grep(parameter.patternText,
                                parameter.pathList.toArray(new Path[parameter.pathList.size()]));
        long runningTime = (System.currentTimeMillis() - startTime);

        long totalMemory = Runtime.getRuntime().totalMemory();
        long usedMemory = totalMemory - Runtime.getRuntime().freeMemory();

        if (parameter.summary) {
            System.out.println("--- Result -------------------------------------------------------------");
        }

        summary.resultFileList.stream()
                .flatMap(rf -> rf.resultList.stream())
                .forEach(r ->
                        System.out.println("["
                                + r.filePath + "]["
                                + r.sheetName + "]["
                                + r.cellAddress + "] "
                                + r.cellValue));

        if (parameter.summary) {
            double mib = 1024 * 1024;
            System.out.println();
            System.out.println("--- Result Summary (File) ----------------------------------------------");
            final String fmt = "[%" + summary.resultFileList.stream()
                    .mapToInt(rf -> String.valueOf(rf.resultList.size()).length()).max().orElse(0) + "d]";
            summary.resultFileList
                    .forEach(rf -> System.out.println(String.format(fmt, rf.resultList.size()) + " : " + rf.file));
            System.out.println();
            System.out.println("--- Result Summary (Total) ---------------------------------------------");
            System.out.println("number of files (matches/total) : "
                    + summary.matchFileCount() + "/" + summary.targetFileCount());
            System.out.println("number of matches               : " + summary.allResultList().size());
            System.out.println("running time                    : " + (runningTime / 1000D) + "s");
            System.out.println("memory (used/total)             : "
                    + String.format("%.1fMB/%.1fMB", usedMemory / mib, totalMemory / mib));
        }
    }

    /**
     * ヘルプを表示します。
     *
     * @param parser コマンドライン引数のパーサ
     * @param out    出力ストリーム
     */
    private void help(CmdLineParser parser, PrintStream out) {
        out.print("Usage: " + COMMAND);
        parser.printSingleLineUsage(out);
        out.println();
        out.println("Search for "
                + CommandLineParameter.META_VAR_PATTERN_TEXT
                + " in each Excel "
                + CommandLineParameter.META_VAR_PATH_LIST
                + ". ("
                + Stream.of(GrepExcel.availableExtensions()).map(ext -> "." + ext).collect(Collectors.joining(", "))
                + ")");
        out.println();
        out.println("Options:");
        parser.printUsage(out);
    }

    /**
     * バージョンを表示します。
     */
    private void version() {
        String version = "x.x.x";
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(ClassLoader.getSystemResourceAsStream("version.txt")))) {
            version = br.lines().findFirst().orElse(version);
        } catch (IOException ignored) {
        }
        System.out.println(COMMAND + " version " + version
                + " (" + "Java version " + System.getProperty("java.version") + ")");
    }

}
