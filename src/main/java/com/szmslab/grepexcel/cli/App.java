/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel.cli;

/**
 * アプリケーションのスタートアップクラスです。
 *
 * @author szmslab
 */
public class App {

    /**
     * メインメソッドです。
     *
     * @param args コマンドライン引数。
     */
    public static void main(String[] args) {
        System.exit(new CommandLineRunner().run(args));
    }

}
