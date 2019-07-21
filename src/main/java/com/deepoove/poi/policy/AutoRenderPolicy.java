/*
 * Copyright 2014-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.deepoove.poi.policy;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.StyleUtils;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import sun.awt.Symbol;

/**
 * @author Sayi
 * @version
 */
public class AutoRenderPolicy extends AbstractRenderPolicy<Object> {

    @Override
    public void doRender(RunTemplate runTemplate, Object renderData, XWPFTemplate template)
            throws Exception {
		throw new RuntimeException("not support");
    }
}
