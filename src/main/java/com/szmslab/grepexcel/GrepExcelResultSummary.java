/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.stream.Collectors;

/**
 * Excelファイル検索結果（全体）を保持するクラスです。
 *
 * @author szmslab
 */
public class GrepExcelResultSummary {

    /**
     * Excelファイル検索結果（ファイル単位）のリスト。
     */
    public final List<GrepExcelResultFile> resultFileList;

    /**
     * コンストラクタです。
     *
     * @param parallel 検索処理を並列実行するかどうか。
     */
    public GrepExcelResultSummary(boolean parallel) {
        this.resultFileList = parallel ? new CopyOnWriteArrayList<>() : new ArrayList<>();
    }

    /**
     * 検索対象のファイル数を取得します。
     *
     * @return 検索対象のファイル数
     */
    public int targetFileCount() {
        return resultFileList.size();
    }

    /**
     * 検索結果が存在するファイル数を取得します。
     *
     * @return 検索結果が存在するファイル数
     */
    public int matchFileCount() {
        return resultFileList.stream().mapToInt(rf -> rf.resultList.size() > 0 ? 1 : 0).sum();
    }

    /**
     * 全てのExcelファイル検索結果（セル単位）のリストを取得します。
     *
     * @return 全てのExcelファイル検索結果（セル単位）のリスト
     */
    public List<GrepExcelResult> allResultList() {
        return resultFileList.stream().flatMap(rf -> rf.resultList.stream()).collect(Collectors.toList());
    }

    @Override
    public String toString() {
        return "{" +
                "resultFileList=" + resultFileList +
                "}";
    }

}
