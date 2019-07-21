/*
 * Copyright 2014-2019 the original author or authors.
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
package com.deepoove.poi.render;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.deepoove.poi.config.GramerSymbol;
import com.deepoove.poi.policy.AutoRenderPolicy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.DocxRenderPolicy;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.util.ObjectUtils;

import javax.swing.text.AbstractDocument;

/**
 * 渲染器，支持表达式计算接口RenderDataCompute的扩展
 *
 * @author Sayi
 * @since 1.5.0
 */
public class Render {

	private final Logger LOGGER = LoggerFactory.getLogger(Render.class);

	private RenderDataCompute renderDataCompute;

	/**
	 * 默认计算器为ELObjectRenderDataCompute
	 *
	 * @param root
	 */
	public Render(Object root) {
		ObjectUtils.requireNonNull(root, "Data root is null, should be setted first.");
		renderDataCompute = new ELObjectRenderDataCompute(root, false);
	}

	public Render(RenderDataCompute dataCompute) {
		this.renderDataCompute = dataCompute;
	}

	public void render(XWPFTemplate template) {
		ObjectUtils.requireNonNull(template, "Template is null, should be setted first.");
		LOGGER.info("Render the template file start...");

		// 模板
		List<ElementTemplate> elementTemplates = template.getElementTemplates();
		if (null == elementTemplates || elementTemplates.isEmpty()) {
			return;
		}
		// 策略
		RenderPolicy policy = null;

		try {
			int docxCount = 0;
			Map<String, Object> templateContent = new HashMap<String, Object>();
			for (ElementTemplate runTemplate : elementTemplates) {
				Object content = renderDataCompute.compute(runTemplate.getTagName());
				templateContent.put(runTemplate.getSource(), content);
				policy = findPolicy(template.getConfig(), runTemplate, content);

				if (policy instanceof DocxRenderPolicy) {
					docxCount++;
					continue;
				}
				doRender(runTemplate, content, policy, template);
			}

			if (docxCount >= 1) template.reload(template.getXWPFDocument().generate());

			NiceXWPFDocument current = null;
			for (int i = 0; i < docxCount; i++) {
				current = template.getXWPFDocument();
				elementTemplates = template.getElementTemplates();
				if (null == elementTemplates || elementTemplates.isEmpty()) {
					break;
				}

				for (ElementTemplate runTemplate : elementTemplates) {
					Object data = templateContent.get(runTemplate.getSource());
					policy = findPolicy(template.getConfig(), runTemplate, data);
					if (!(policy instanceof DocxRenderPolicy)) {
						continue;
					}
					doRender(runTemplate, data, policy, template);

					// 没有最终合并，继续下一个合并
					if (current == template.getXWPFDocument()) {
						i++;
						continue;
					} else {
						break;
					}
				}
			}
		} catch (Exception e) {
			LOGGER.info("Render the template file failed.");
			throw new RenderException("Render docx failed.", e);
		}
		LOGGER.info("Render the template file successed.");
	}

	private RenderPolicy findPolicy(Configure config, ElementTemplate runTemplate, Object content) {
		RenderPolicy policy;
		policy = config.getPolicy(runTemplate.getTagName(), runTemplate.getSign());
		if (policy instanceof AutoRenderPolicy) {
			if (content == null) {
				policy = config.getPolicy(null, GramerSymbol.TEXT.getSymbol());
			} else {
				policy = config.getPolicy(null, config.guessSign(content.getClass()));
			}
		}

		if (null == policy) {
			throw new RenderException(
				"Cannot find render policy: [" + runTemplate.getTagName() + "]");
		}
		return policy;
	}

	private void doRender(ElementTemplate ele, Object content, RenderPolicy policy, XWPFTemplate template) {
		LOGGER.debug("Start render TemplateName:{}, Sign:{}, policy:{}", ele.getTagName(),
			ele.getSign(), policy.getClass().getSimpleName());
		policy.render(ele, content, template);
	}

}
