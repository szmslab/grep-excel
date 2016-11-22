/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * Excelファイル検索結果（ファイル単位）を保持するクラスです。
 *
 * @author szmslab
 */
public class GrepExcelResultFile {

    /**
     * 検索対象のファイルパス。
     */
    public final Path file;

    /**
     * コンストラクタです。
     *
     * @param file 検索対象のファイルパス
     */
    public GrepExcelResultFile(Path file) {
        this.file = file;
    }

    /**
     * コンストラクタです。
     *
     * @param file       検索対象のファイルパス
     * @param resultList Excelファイル検索結果（セル単位）のリスト
     */
    public GrepExcelResultFile(Path file, List<GrepExcelResult> resultList) {
        this.file = file;
        this.resultList.addAll(resultList);
    }

    /**
     * Excelファイル検索結果（セル単位）のリスト。
     */
    public final List<GrepExcelResult> resultList = new ArrayList<>();

    @Override
    public String toString() {
        return "{" +
                "file=" + file +
                ", resultList=" + resultList +
                "}";
    }

}
