/*
 * Copyright (c) 2016 szmslab
 *
 * This software is released under the MIT License.
 * http://opensource.org/licenses/mit-license.php
 */
package com.szmslab.grepexcel.cli;

import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;
import org.kohsuke.args4j.OptionDef;
import org.kohsuke.args4j.spi.PathOptionHandler;
import org.kohsuke.args4j.spi.Setter;

import java.io.FileNotFoundException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * 存在するファイル・ディレクトリを{@link Path}にマッピングするオプションハンドラクラスです。
 *
 * @author szmslab
 */
public class ExistingPathOptionHandler extends PathOptionHandler {

    public ExistingPathOptionHandler(CmdLineParser parser, OptionDef option, Setter<? super Path> setter) {
        super(parser, option, setter);
    }

    @Override
    protected Path parse(String argument) throws NumberFormatException, CmdLineException {
        Path path = super.parse(argument);
        if (!Files.exists(path)) {
            FileNotFoundException e = new FileNotFoundException("No such file or directory \"" + argument + "\"");
            throw new CmdLineException(owner, e.getMessage(), e);
        }
        return path;
    }

}
