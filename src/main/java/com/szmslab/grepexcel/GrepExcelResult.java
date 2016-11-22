/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel;

/**
 * Excelファイル検索結果（セル単位）を保持するクラスです。
 *
 * @author szmslab
 */
public class GrepExcelResult {

    /**
     * ファイルパス。
     */
    public final String filePath;

    /**
     * ワークシート名。
     */
    public final String sheetName;

    /**
     * セルのアドレス。
     */
    public final String cellAddress;

    /**
     * セルの値。
     */
    public final String cellValue;

    /**
     * コンストラクタです。
     *
     * @param filePath    ファイルパス
     * @param sheetName   ワークシート名
     * @param cellAddress セルのアドレス
     * @param cellValue   セルの値
     */
    public GrepExcelResult(String filePath, String sheetName, String cellAddress, String cellValue) {
        this.filePath = filePath;
        this.sheetName = sheetName;
        this.cellAddress = cellAddress;
        this.cellValue = cellValue;
    }

    @Override
    public String toString() {
        return "{" +
                "filePath=" + filePath +
                ", sheetName=" + sheetName +
                ", cellAddress=" + cellAddress +
                ", cellValue=" + cellValue +
                "}";
    }

}
