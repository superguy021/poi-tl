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
package com.deepoove.poi.policy;

import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

/**
 * 抽象策略
 *
 * @author Sayi
 */
public abstract class AbstractRenderPolicy<T> implements RenderPolicy {

	private Class<?> clazz;

	{
		Type type = getClass().getGenericSuperclass();
		if (type instanceof ParameterizedType) {
			ParameterizedType pType = (ParameterizedType) type;
			Type claz = pType.getActualTypeArguments()[0];
			if (claz instanceof Class) {
				this.clazz = (Class<T>) claz;
			}
		}
	}

	/**
	 * 校验data model
	 *
	 * @param data
	 * @return
	 */
	protected boolean validate(T data) {
		return true;
	}

	;

	/**
	 * 校验失败
	 *
	 * @param template
	 * @param runTemplate
	 */
	protected void doValidError(RenderContext context) {
		logger.debug("Validate the data of the element {} error, the data may be null or empty: {}",
			context.getEleTemplate().getSource(), context.getData());
		if (context.getTemplate().getConfig().isNullToBlank()) {
			logger.debug("[config.isNullToBlank == true] clear the element {} from the word file.",
				context.getEleTemplate().getSource());
			clearPlaceholder(context, false);
		} else {
			logger.debug("The element {} Unable to be rendered, nothing to do.",
				context.getEleTemplate().getSource());
		}
	}

	protected void beforeRender(RenderContext context) {
	}

	protected void afterRender(RenderContext context) {
	}

	/**
	 * 执行模板渲染
	 *
	 * @param runTemplate 文档模板
	 * @param data        数据模型
	 * @param template    文档对象
	 * @throws Exception
	 */
	public abstract void doRender(RunTemplate runTemplate, T data, XWPFTemplate template)
		throws Exception;

	/*
	 * 骨架 (non-Javadoc)
	 *
	 * @see
	 * com.deepoove.poi.policy.RenderPolicy#render(com.deepoove.poi.template.
	 * ElementTemplate, java.lang.Object, com.deepoove.poi.XWPFTemplate)
	 */
	@SuppressWarnings("unchecked")
	@Override
	public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
		RunTemplate runTemplate = (RunTemplate) eleTemplate;

		// type safe
		T model = null;
		try {
			model = (T) data;
		} catch (ClassCastException e) {
			throw new RenderException(
				"Error Render Data format for template: " + eleTemplate.getSource(), e);
		}

		// validate
		RenderContext context = new RenderContext(eleTemplate, data, template);
		if (null == model || !validate(model)) {
			doValidError(context);
			return;
		}

		// do render
		try {
			beforeRender(context);
			doRender(runTemplate, model, template);
			afterRender(context);
		} catch (Exception e) {
			doRenderException(runTemplate, model, e);
		}

	}

	/**
	 * 发生异常
	 *
	 * @param runTemplate
	 * @param data
	 * @param e
	 */
	protected void doRenderException(RunTemplate runTemplate, T data, Exception e) {
		throw new RenderException("Render template:" + runTemplate + " error", e);
	}

	/**
	 * 继承这个方法，实现自定义清空标签的方案
	 *
	 * @param context
	 * @param clearParagraph 是否清空占位的段落
	 */
	protected void clearPlaceholder(RenderContext context, boolean clearParagraph) {
		XWPFRun run = ((RunTemplate) context.getEleTemplate()).getRun();
		run.setText("", 0);
		IRunBody parent = run.getParent();
		if (clearParagraph && (parent instanceof XWPFParagraph)) {
			NiceXWPFDocument doc = context.getTemplate().getXWPFDocument();
			int posOfParagraph = doc.getPosOfParagraph((XWPFParagraph) parent);
			doc.removeBodyElement(posOfParagraph);
		}
	}

	public Class workDataType() {
		return clazz;
	}
}
