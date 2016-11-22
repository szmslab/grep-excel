/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel;

import com.github.mygreen.cellformatter.POICellFormatter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.FileVisitOption;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Excelファイル内の文字列を検索するクラスです。
 *
 * @author szmslab
 */
public class GrepExcel {

    /**
     * ロガー。
     */
    private static final Logger LOG = LoggerFactory.getLogger(GrepExcel.class);

    /**
     * Excelファイルの拡張子。
     */
    private static final String[] EXTENSIONS = {"xls", "xlsx", "xlsm", "xlt", "xltx"};

    /**
     * セル値取得フォーマッタ。
     */
    private final POICellFormatter formatter = new POICellFormatter();

    /**
     * 大文字・小文字を区別しないかどうか。
     */
    private boolean ignoreCase;

    /**
     * リテラル構文解析を有効にするかどうか。
     */
    private boolean literal;

    /**
     * 数式セルの計算結果を検索するかどうか。
     */
    private boolean formulaResult;

    /**
     * ディレクトリを再帰的に検索するかどうか。
     */
    private boolean recursive;

    /**
     * 検索処理を並列実行するかどうか。
     */
    private boolean parallel;

    /**
     * 利用可能なExcelファイルの拡張子を取得します。
     *
     * @return 利用可能なExcelファイルの拡張子
     */
    public static String[] availableExtensions() {
        return EXTENSIONS.clone();
    }

    /**
     * 大文字・小文字を区別しないかどうかを設定します。
     *
     * @param ignoreCase 大文字・小文字を区別しない場合は {@code true}
     * @return 自身のインスタンス
     */
    public GrepExcel ignoreCase(boolean ignoreCase) {
        this.ignoreCase = ignoreCase;
        return this;
    }

    /**
     * 大文字・小文字を区別しないかどうかを取得します。
     *
     * @return 大文字・小文字を区別しない場合は {@code true}
     */
    public boolean ignoreCase() {
        return ignoreCase;
    }

    /**
     * リテラル構文解析を有効にするかどうかを設定します。
     *
     * @param literal リテラル構文解析を有効にする場合は {@code true}
     * @return 自身のインスタンス
     */
    public GrepExcel literal(boolean literal) {
        this.literal = literal;
        return this;
    }

    /**
     * リテラル構文解析を有効にするかどうかを取得します。
     *
     * @return リテラル構文解析を有効にする場合は {@code true}
     */
    public boolean literal() {
        return literal;
    }

    /**
     * 数式セルの計算結果を検索するかどうかを設定します。
     *
     * @param formulaResult 数式セルの計算結果を検索する場合は {@code true}
     * @return 自身のインスタンス
     */
    public GrepExcel formulaResult(boolean formulaResult) {
        this.formulaResult = formulaResult;
        return this;
    }

    /**
     * 数式セルの計算結果を検索するかどうかを取得します。
     *
     * @return 数式セルの計算結果を検索する場合は {@code true}
     */
    public boolean formulaResult() {
        return formulaResult;
    }

    /**
     * ディレクトリを再帰的に検索するかどうかを設定します。
     *
     * @param recursive ディレクトリを再帰的に検索する場合は {@code true}
     * @return 自身のインスタンス
     */
    public GrepExcel recursive(boolean recursive) {
        this.recursive = recursive;
        return this;
    }

    /**
     * ディレクトリを再帰的に検索するかどうかを取得します。
     *
     * @return ディレクトリを再帰的に検索する場合は {@code true}
     */
    public boolean recursive() {
        return recursive;
    }

    /**
     * 検索処理を並列実行するかどうかを設定します。
     *
     * @param parallel 検索処理を並列実行する場合は {@code true}
     * @return 自身のインスタンス
     */
    public GrepExcel parallel(boolean parallel) {
        this.parallel = parallel;
        return this;
    }

    /**
     * 検索処理を並列実行するかどうかを取得します。
     *
     * @return 検索処理を並列実行する場合は {@code true}
     */
    public boolean parallel() {
        return parallel;
    }

    /**
     * Excelファイル内の文字列を検索します。
     *
     * @param patternText 検索パターン
     * @param paths       検索対象のファイル・ディレクトリパス
     * @return Excelファイル検索結果（全体）
     */
    public GrepExcelResultSummary grep(String patternText, Path... paths) {
        LOG.debug("fields: {}", this);
        LOG.debug("patternText: {}", patternText);
        LOG.debug("paths: {}", Arrays.toString(paths));

        Pattern pattern = Pattern.compile(patternText,
                Pattern.MULTILINE
                        | Pattern.DOTALL
                        | (ignoreCase ? Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE : 0x00)
                        | (literal ? Pattern.LITERAL : 0x00));

        GrepExcelResultSummary summary = new GrepExcelResultSummary(parallel);
        Stream<Path> stream = Stream.of(getExcelFiles(paths));
        if (parallel) {
            stream = stream.parallel();
        }
        stream.forEach(file -> summary.resultFileList.add(new GrepExcelResultFile(file, grep(pattern, file))));
        summary.resultFileList.sort((o1, o2) -> o1.file.compareTo(o2.file));

        return summary;
    }

    /**
     * Excelファイル内の文字列を検索します。
     *
     * @param pattern コンパイルされた検索パターン
     * @param file    検索対象のファイルパス
     * @return Excelファイル検索結果（ファイル）
     */
    private List<GrepExcelResult> grep(Pattern pattern, Path file) {
        try (Workbook book = WorkbookFactory.create(Files.newInputStream(file))) {
            return grep(pattern, file, book);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(file.toString(), e);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }

    /**
     * Excelファイル内の文字列を検索します。
     *
     * @param pattern コンパイルされた検索パターン
     * @param file    検索対象のファイルパス
     * @param book    ワークブック
     * @return Excelファイル検索結果（ワークブック）
     */
    private List<GrepExcelResult> grep(Pattern pattern, Path file, Workbook book) {
        return toStream(book.sheetIterator(), book.getNumberOfSheets())
                .flatMap(sheet -> grep(pattern, file, sheet).stream())
                .collect(Collectors.toList());
    }

    /**
     * Excelファイル内の文字列を検索します。
     *
     * @param pattern コンパイルされた検索パターン
     * @param file    検索対象のファイルパス
     * @param sheet   ワークシート
     * @return Excelファイル検索結果（ワークシート）
     */
    private List<GrepExcelResult> grep(Pattern pattern, Path file, Sheet sheet) {
        return toStream(sheet.rowIterator())
                .flatMap(row -> grep(pattern, file, sheet, row).stream())
                .collect(Collectors.toList());
    }

    /**
     * Excelファイル内の文字列を検索します。
     *
     * @param pattern コンパイルされた検索パターン
     * @param file    検索対象のファイルパス
     * @param sheet   ワークシート
     * @param row     行
     * @return Excelファイル検索結果（行）
     */
    private List<GrepExcelResult> grep(Pattern pattern, Path file, Sheet sheet, Row row) {
        List<GrepExcelResult> list = new ArrayList<>();
        for (Iterator<Cell> itr = row.cellIterator(); itr.hasNext(); ) {
            Cell cell = itr.next();

            if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                continue;
            }

            String cellAddress = cell.getAddress().toString();
            String cellValue = toStringValue(cell);

            if (LOG.isDebugEnabled()) {
                LOG.debug("file: {}, sheet: {}, cell: {}, value: {}",
                        file.getFileName(), sheet.getSheetName(), cellAddress, cellValue);
            }

            if (pattern.matcher(cellValue).find()) {
                list.add(new GrepExcelResult(
                        file.toString(), sheet.getSheetName(), cellAddress, cellValue));
            }
        }
        return list;
    }

    /**
     * セルの値を文字列で取得します。
     *
     * @param cell セル
     * @return セルの文字列値
     */
    private String toStringValue(Cell cell) {
        if (cell.getCellType() == Cell.CELL_TYPE_FORMULA && !formulaResult) {
            return "=" + cell.getCellFormula();
        }

        try {
            return formatter.formatAsString(cell);
        } catch (Exception e) {
            return cell.toString();
        }
    }

    /**
     * 指定したファイル・ディレクトリパスから、Excelファイルのパスを取得します。
     *
     * @param paths ファイル・ディレクトリパス
     * @return Excelファイルのパス
     */
    private Path[] getExcelFiles(Path... paths) {
        return Stream.of(paths)
                .flatMap(p -> {
                    try {
                        return Files.walk(p, recursive ? Integer.MAX_VALUE : 1, FileVisitOption.FOLLOW_LINKS);
                    } catch (IOException e) {
                        throw new UncheckedIOException(e);
                    }
                })
                .filter(this::isExcelFile)
                .map(path -> path.toAbsolutePath().normalize())
                .toArray(Path[]::new);
    }

    /**
     * Excelファイルかどうかを取得します。
     *
     * @param path ファイル・ディレクトリパス
     * @return Excelファイルの場合は {@code true}
     */
    private boolean isExcelFile(Path path) {
        if (!Files.isRegularFile(path)) {
            return false;
        }
        String fileName = path.toAbsolutePath().normalize().getFileName().toString();
        int idx = fileName.lastIndexOf(".");
        return idx >= 0 && Arrays.asList(EXTENSIONS).contains(fileName.substring(idx + 1));
    }

    /**
     * イテレータをストリームに変換して返します。
     *
     * @param iterator 変換対象のイテレータ
     * @param <T>      イテレータが返す型
     * @return 変換されたストリーム
     */
    private <T> Stream<T> toStream(Iterator<? extends T> iterator) {
        return toStream(iterator, -1);
    }

    /**
     * イテレータをストリームに変換して返します。
     *
     * @param iterator 変換対象のイテレータ
     * @param size     イテレータのサイズ
     * @param <T>      イテレータが返す型
     * @return 変換されたストリーム
     */
    private <T> Stream<T> toStream(Iterator<? extends T> iterator, long size) {
        int characteristics = Spliterator.SIZED | Spliterator.ORDERED;
        Spliterator<T> spliterator = size >= 0
                ? Spliterators.spliterator(iterator, size, characteristics)
                : Spliterators.spliteratorUnknownSize(iterator, characteristics);
        return StreamSupport.stream(spliterator, false);
    }

    @Override
    public String toString() {
        return "{" +
                "formatter=" + formatter +
                ", ignoreCase=" + ignoreCase +
                ", literal=" + literal +
                ", formulaResult=" + formulaResult +
                ", recursive=" + recursive +
                ", parallel=" + parallel +
                "}";
    }

}
