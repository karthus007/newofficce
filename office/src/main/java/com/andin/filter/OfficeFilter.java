package com.andin.filter;

import java.io.IOException;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.annotation.WebFilter;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Component;

import com.alibaba.fastjson.JSONObject;
import com.andin.utils.CommonUtil;
import com.andin.utils.ConstantUtil;

@Component
@WebFilter(filterName = "filter")
public class OfficeFilter implements Filter{

	@Override
	public void doFilter(ServletRequest request, ServletResponse response, FilterChain chain)
			throws IOException, ServletException {
		HttpServletRequest req = (HttpServletRequest) request;
		HttpServletResponse resp = (HttpServletResponse) response;
		String uri = req.getRequestURI();
		if(uri.contains("/license") || uri.contains(".")) {
			chain.doFilter(request, response);			
		}else {
			if(CommonUtil.LICENSE_STATUS) {
				chain.doFilter(request, response);			
			}else {
				writeAjaxResponse(resp);				
			}
		}
		
	}
	
	
	/**
	 * 返回错误响应
	 * @param resp
	 * @param description
	 * @throws IOException
	 */
	private void writeAjaxResponse(HttpServletResponse resp) throws IOException {
        JSONObject result = new JSONObject();
        result.put(ConstantUtil.RESULT_CODE, ConstantUtil.LICENSE_NOAUTH_ERROR_CODE);
        result.put(ConstantUtil.RESULT_MSG, ConstantUtil.LICENSE_NOAUTH_ERROR_MSG);
        resp.setStatus(200);
        resp.setContentType(ConstantUtil.APPLICATION_JSON_UTF_8);
        resp.getWriter().append(result.toJSONString());
	}

}
